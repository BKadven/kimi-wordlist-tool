"""一次性回填工具：为已有单词本归档生成 wordbook.json 并重建 wordbooks-manifest.json。

用途（每日单词复习邮件系统项目书 · 阶段 A）：
- 优先使用本地留存的 ``*_结构化数据_*.json``（与线上内容完全一致的原始结构化数据）；
- 找不到原始 JSON 的归档，从手机版 index.html 逆向解析还原结构化数据，
  并输出逐条保真度比对报告供人工审阅；
- 默认 dry-run，只产出报告，不写任何文件；
- ``--write`` 时才写入 wordbook.json / wordbooks-manifest.json；
- 本工具绝不执行 git add / commit / push，发布由人工确认后单独进行。

用法（在 源代码 目录下）：

    .venv\\Scripts\\python.exe tools\\backfill_wordbook_json.py
    .venv\\Scripts\\python.exe tools\\backfill_wordbook_json.py --write
"""

from __future__ import annotations

import argparse
import difflib
import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

from lxml import html as lxml_html

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from wordlist_engine import (  # noqa: E402
    build_wordbook_document,
    compute_content_sha256,
    count_wordbook_data,
    sanitize_url_part,
    update_wordbooks_manifest,
    value_text,
)

DEFAULT_WORDBOOKS_DIR = (
    Path.home() / "Documents" / "All_Program" / "personal-homepage"
    / "notes" / "ielts" / "wordbooks"
)
DEFAULT_JSON_DIR = (
    Path.home() / "Desktop" / "总站" / "本地中继站"
    / "【常用资料】" / "雅思单词本"
)
DEFAULT_BASE_URL = "https://xxxcillian.org"


# ---------------------------------------------------------------- 工具函数

def normalize_text(value: str) -> str:
    """比对用归一化：压缩全部空白。"""
    return re.sub(r"\s+", "", value or "")


def similarity(left: str, right: str) -> float:
    left_n = normalize_text(left)
    right_n = normalize_text(right)
    if not left_n and not right_n:
        return 1.0
    return difflib.SequenceMatcher(None, left_n, right_n).ratio()


def read_meta(archive_dir: Path) -> dict:
    meta_path = archive_dir / "meta.json"
    try:
        meta = json.loads(meta_path.read_text(encoding="utf-8"))
        return meta if isinstance(meta, dict) else {}
    except (json.JSONDecodeError, OSError):
        return {}


# ------------------------------------------------- 结构化 JSON 精确还原路径

def index_structured_jsons(json_dir: Path) -> dict[str, Path]:
    """按 sanitize_url_part(文件主干) 建立索引，供归档目录名匹配。"""
    index: dict[str, Path] = {}
    if not json_dir.is_dir():
        return index
    for path in sorted(json_dir.glob("*结构化数据*.json")):
        stem = path.name.split("_结构化数据_")[0]
        index[sanitize_url_part(stem)] = path
    return index


def find_structured_json(archive_id: str, json_index: dict[str, Path]) -> Path | None:
    for url_part, path in json_index.items():
        if url_part and url_part in archive_id:
            return path
    return None


# --------------------------------------------------- HTML 逆向解析还原路径

def parse_meta_span(text: str) -> tuple[str, str]:
    """解析词条头部 meta（如 "n. · B"），返回 (pos, frequency)。"""
    parts = [part.strip() for part in value_text(text).split("·")]
    pos = parts[0] if parts else ""
    frequency = parts[1] if len(parts) > 1 else ""
    return pos, frequency


def parse_definition_blocks(word_el) -> str:
    """把渲染后的释义区块逆向拼回单一 definition 字符串。

    正向渲染时 definition 被按 " / " 和【…】批注拆开渲染；这里按文档顺序
    收集 definition-point / phrase-card 条目 / note-point 的文本并重新拼接。
    """
    single = word_el.find('.//p[@class="definition"]')
    detail_list = word_el.find('.//div[@class="detail-list"]')
    if detail_list is None:
        return value_text(single.text_content() if single is not None else "")

    chunks: list[str] = []
    for child in detail_list:
        classes = (child.get("class") or "").split()
        if "phrase-card" in classes:
            for li in child.findall('.//ul[@class="phrase-list"]/li'):
                text = value_text(li.text_content())
                if text:
                    chunks.append(text)
        else:
            text = value_text(child.text_content())
            if text:
                chunks.append(text)
    return " / ".join(chunks)


def parse_word_item(word_el, group: str) -> tuple[dict, list[str]]:
    warnings: list[str] = []
    header = word_el.find("./header")
    if header is not None:
        headword = value_text(header.findtext("h3"))
        pos, frequency = parse_meta_span(
            header.findtext('span[@class="meta"]') or ""
        )
    else:
        # 正向渲染规则：headword 与 group 相同则不输出 header
        headword, pos, frequency = group, "", ""
        warnings.append(f"词条“{group}”缺少逐词词性/频度（HTML 未渲染头部）")

    collocations = [
        value_text(span.text_content())
        for span in word_el.findall('.//div[@class="chips"]/span')
        if value_text(span.text_content())
    ]

    usage_note = ""
    usage_el = word_el.find('.//p[@class="usage"]')
    if usage_el is not None:
        usage_note = re.sub(r"^使用提醒[：:]\s*", "", value_text(usage_el.text_content()))

    item = {
        "headword": headword or group,
        "pos": pos,
        "frequency": frequency,
        "definition": parse_definition_blocks(word_el),
        "collocations": collocations,
        "usage_note": usage_note,
    }
    return item, warnings


def parse_family_additions(entry_el) -> list[dict]:
    block = entry_el.find('.//div[@class="note-block"][h4="词族补充"]')
    if block is None:
        return []
    additions: list[dict] = []
    for li in block.findall("./ul/li"):
        headword = value_text(li.findtext("strong"))
        meta = value_text(li.findtext("span"))
        pos, frequency = parse_meta_span(meta)
        full = value_text(li.text_content())
        definition = ""
        match = re.search(r"[：:]\s*(.+)$", full, flags=re.DOTALL)
        if match:
            definition = match.group(1).strip()
        if headword:
            additions.append(
                {
                    "headword": headword,
                    "pos": pos,
                    "frequency": frequency,
                    "definition": definition,
                }
            )
    return additions


def parse_note_block(entry_el, title: str) -> str:
    block = entry_el.find(f'.//div[@class="note-block"][h4="{title}"]/p')
    if block is None:
        return ""
    return value_text(block.text_content())


def parse_corrections(root) -> list[dict]:
    corrections: list[dict] = []
    for li in root.findall('.//section[@class="corrections"]/ul/li'):
        strongs = li.findall("strong")
        original = value_text(strongs[0].text_content()) if len(strongs) >= 1 else ""
        corrected = value_text(strongs[1].text_content()) if len(strongs) >= 2 else ""
        note = ""
        match = re.search(r"（(.+?)）\s*$", value_text(li.text_content()))
        if match:
            note = match.group(1).strip()
        if original or corrected:
            corrections.append(
                {"original": original, "corrected": corrected, "note": note}
            )
    return corrections


def parse_mobile_html(index_path: Path) -> tuple[dict, list[str], list[str]]:
    """逆向解析手机版 HTML，返回 (结构化数据, 警告列表, 原文条目文本列表)。"""
    warnings: list[str] = []
    original_entry_texts: list[str] = []
    root = lxml_html.fromstring(index_path.read_bytes())

    subtitle_el = root.find('.//p[@class="subtitle"]')
    subtitle = value_text(subtitle_el.text_content()) if subtitle_el is not None else ""

    sections: list[dict] = []
    for details_el in root.findall('.//details[@class="section"]'):
        title = value_text(details_el.findtext("summary"))
        entries: list[dict] = []
        for entry_el in details_el.findall('./article[@data-entry]'):
            original_entry_texts.append(entry_el.text_content())
            group = value_text(entry_el.findtext('.//div[@class="entry-title"]/h2'))
            frequency_range = value_text(
                entry_el.findtext('.//div[@class="entry-title"]/span')
            )
            items: list[dict] = []
            for word_el in entry_el.findall('./div[@class="words"]/article[@class="word"]'):
                item, item_warnings = parse_word_item(word_el, group)
                items.append(item)
                warnings.extend(item_warnings)
            if not items:
                items.append(
                    {
                        "headword": group,
                        "pos": "",
                        "frequency": "",
                        "definition": "",
                        "collocations": [],
                        "usage_note": "",
                    }
                )
                warnings.append(f"词条“{group}”在 HTML 中没有可解析的逐词内容")
            entries.append(
                {
                    "group": group,
                    "frequency_range": frequency_range,
                    "items": items,
                    "family_additions": parse_family_additions(entry_el),
                    "difference": parse_note_block(entry_el, "区别"),
                    "group_note": parse_note_block(entry_el, "补充说明"),
                }
            )
        if title and entries:
            sections.append({"title": title, "entries": entries})

    data = {
        "subtitle": subtitle,
        "corrections": parse_corrections(root),
        "sections": sections,
    }
    return data, warnings, original_entry_texts


# --------------------------------------------------------------- 保真度比对

def rerendered_entry_texts(data: dict) -> list[str]:
    """用当前渲染器重渲染还原数据，抽取每个词条的纯文本用于比对。"""
    from wordlist_engine import render_mobile_entry

    texts: list[str] = []
    for section in data.get("sections", []):
        for entry in section.get("entries", []):
            fragment = lxml_html.fromstring(render_mobile_entry(entry))
            texts.append(fragment.text_content())
    return texts


@dataclass
class BookReport:
    archive_id: str
    source: str = ""                 # exact-json / parsed-html
    status: str = "pending"          # ok / dry-run / error / skipped
    entries: int = 0
    words: int = 0
    meta_entries: int | None = None
    meta_words: int | None = None
    min_similarity: float = 1.0
    avg_similarity: float = 1.0
    warnings: list[str] = field(default_factory=list)
    error: str = ""
    wordbook: dict | None = None


def build_report_line(report: BookReport) -> str:
    sim = f"{report.avg_similarity:.3f}/{report.min_similarity:.3f}"
    meta = ""
    if report.meta_entries is not None:
        flag = "✓" if (report.meta_entries == report.entries and report.meta_words == report.words) else "✗"
        meta = f" meta{flag}({report.meta_entries}/{report.meta_words})"
    warn = f" ⚠{len(report.warnings)}" if report.warnings else ""
    err = f" ERROR:{report.error}" if report.error else ""
    return (
        f"| {report.archive_id} | {report.source} | {report.status} | "
        f"{report.entries}/{report.words}{meta} | {sim}{warn}{err} |"
    )


def process_archive(
    archive_dir: Path,
    json_index: dict[str, Path],
    *,
    write: bool,
) -> BookReport:
    archive_id = archive_dir.name
    report = BookReport(archive_id=archive_id)
    meta = read_meta(archive_dir)
    report.meta_entries = meta.get("entry_count")
    report.meta_words = meta.get("word_count")

    wordbook_json_path = archive_dir / "wordbook.json"
    if wordbook_json_path.is_file():
        report.status = "skipped(已存在)"
        report.source = "-"
        return report

    source_file = value_text(meta.get("source_file")) or f"{archive_id}.docx"
    generated_at = value_text(meta.get("generated_at")) or None
    source_path = Path(source_file)

    try:
        json_path = find_structured_json(archive_id, json_index)
        if json_path is not None:
            report.source = f"exact-json({json_path.name})"
            data = json.loads(json_path.read_text(encoding="utf-8"))
            if not isinstance(data, dict) or not isinstance(data.get("sections"), list):
                raise ValueError(f"结构化 JSON 缺少 sections：{json_path}")
        else:
            report.source = "parsed-html"
            index_path = archive_dir / "index.html"
            data, warnings, original_texts = parse_mobile_html(index_path)
            report.warnings.extend(warnings)
            if not data["sections"]:
                raise ValueError("index.html 中没有解析到任何章节")

            # 保真度：还原数据重渲染 vs 原始 HTML 纯文本
            new_texts = rerendered_entry_texts(data)
            scores: list[float] = []
            for old_text, new_text in zip(original_texts, new_texts):
                scores.append(similarity(old_text, new_text))
            if len(original_texts) != len(new_texts):
                report.warnings.append(
                    f"词条数量不一致：原 HTML {len(original_texts)} 条，还原 {len(new_texts)} 条"
                )
            if scores:
                report.min_similarity = min(scores)
                report.avg_similarity = sum(scores) / len(scores)
                for entry_el_score, (section, entry) in zip(
                    scores,
                    [
                        (s, e)
                        for s in data["sections"]
                        for e in s["entries"]
                    ],
                ):
                    if entry_el_score < 0.98:
                        report.warnings.append(
                            f"词条“{entry.get('group')}”保真度 {entry_el_score:.3f} < 0.98"
                        )

        report.entries, report.words = count_wordbook_data(data)
        wordbook = build_wordbook_document(
            data, source_path, archive_id, generated_at=generated_at
        )
        # 保留 meta.json 中的标题（与收录页展示一致）
        if value_text(meta.get("title")):
            wordbook["title"] = value_text(meta.get("title"))
        report.wordbook = wordbook

        if write:
            wordbook_json_path.write_text(
                json.dumps(wordbook, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        report.status = "written" if write else "dry-run"
    except Exception as exc:  # noqa: BLE001 - 单本失败不阻断整体回填
        report.status = "error"
        report.error = str(exc)
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description="回填已有单词本的 wordbook.json")
    parser.add_argument("--wordbooks-dir", type=Path, default=DEFAULT_WORDBOOKS_DIR)
    parser.add_argument("--json-dir", type=Path, default=DEFAULT_JSON_DIR)
    parser.add_argument("--base-url", default=DEFAULT_BASE_URL)
    parser.add_argument("--write", action="store_true", help="实际写入 wordbook.json 与 manifest")
    parser.add_argument(
        "--report",
        type=Path,
        default=Path(__file__).resolve().parent.parent.parent
        / ".workbuddy" / "backfill_report.md",
    )
    args = parser.parse_args()

    wordbooks_dir: Path = args.wordbooks_dir
    if not wordbooks_dir.is_dir():
        print(f"找不到单词本归档目录：{wordbooks_dir}")
        return 1

    json_index = index_structured_jsons(args.json_dir)
    archives = sorted(
        (child for child in wordbooks_dir.iterdir() if child.is_dir()),
        key=lambda path: path.name,
    )

    reports: list[BookReport] = []
    for archive_dir in archives:
        reports.append(process_archive(archive_dir, json_index, write=args.write))

    if args.write:
        written = [r for r in reports if r.wordbook and r.status == "written"]
        # 逐本 upsert，保证 manifest 与磁盘一致
        for report in written:
            update_wordbooks_manifest(wordbooks_dir, report.wordbook, base_url=args.base_url)

    exact = sum(1 for r in reports if r.source.startswith("exact-json"))
    parsed = sum(1 for r in reports if r.source == "parsed-html")
    errors = [r for r in reports if r.status == "error"]
    skipped = sum(1 for r in reports if r.status.startswith("skipped"))
    low_fidelity = [
        r for r in reports if r.source == "parsed-html" and r.min_similarity < 0.98
    ]

    lines = [
        "# wordbook.json 回填报告",
        "",
        f"- 模式：{'WRITE（已写盘）' if args.write else 'DRY-RUN（未写盘）'}",
        f"- 归档目录：{wordbooks_dir}",
        f"- 归档总数：{len(archives)}；精确还原 {exact} 本；HTML 逆向 {parsed} 本；"
        f"跳过 {skipped} 本；失败 {len(errors)} 本",
        f"- 保真度 < 0.98 的 HTML 逆向归档：{len(low_fidelity)} 本",
        "",
        "| 归档 | 来源 | 状态 | 词条/逐词(meta比对) | 保真度 平均/最低 |",
        "| --- | --- | --- | --- | --- |",
    ]
    lines.extend(build_report_line(r) for r in reports)
    lines.append("")
    warn_reports = [r for r in reports if r.warnings]
    if warn_reports:
        lines.append("## 警告明细")
        lines.append("")
        for r in warn_reports:
            lines.append(f"### {r.archive_id}")
            for warning in r.warnings:
                lines.append(f"- {warning}")
            lines.append("")

    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text("\n".join(lines), encoding="utf-8")
    print(f"报告已写入：{args.report}")
    print(f"总计 {len(archives)} 本；精确 {exact}；HTML 逆向 {parsed}；失败 {len(errors)}")
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
