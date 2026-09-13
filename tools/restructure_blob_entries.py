# -*- coding: utf-8 -*-
"""一次性修复工具：把旧归档 wordbook.json 里的「糊墙 blob 词条」重组成结构化逐词数据。

背景：25 本旧书的 wordbook.json 是从手机版 HTML 逆向解析来的，其中 13 本
共 355 个词条的 items 只有一个「超级 item」，definition 里塞满了整段文本：
    betray (v.)｜A：背叛；出卖… / betray one's country / 【使用提醒】… / give away (phr. v.)｜A：…
邮件系统按结构化字段渲染，这种 blob 只能整段糊出去，排版失控。

本工具把 blob 拆回结构化字段：
    items[]:        headword / pos / frequency / definition / collocations[] / usage_note
    entry 级:       difference（【区别】）/ group_note（【补充说明】）/ family_additions（【词族补充】）

安全性设计（无损保证）：
1. blob 按 " / " 切分，所有片段必须被完整分类，任一片段无法归类 → 该词条保持原样并报告；
2. 词条头用整段 fullmatch 正则提取，字段全部来自原文逐字切片，不重排、不改写；
3. 【词族补充】内部解析后做「剩余字符白名单」校验（只允许分隔符残留），否则该词条保持原样；
4. 默认 dry-run 输出报告，--write 才写盘；写盘前重算 content_sha256 与 entry/word_count，
   并同步更新 wordbooks-manifest.json 对应行。

用法：
    python tools/restructure_blob_entries.py            # dry-run，输出报告
    python tools/restructure_blob_entries.py --write    # 实际写盘
"""

from __future__ import annotations

import hashlib
import json
import re
import sys
from pathlib import Path

HOMEPAGE_WORDBOOKS = Path(
    r"C:\Users\cilli\Documents\All_Program\personal-homepage\notes\ielts\wordbooks"
)
REPORT_PATH = Path(__file__).resolve().parents[2] / ".workbuddy" / "restructure_report.md"

SEGMENT_SEP = " / "
# 词条头：headword (pos)｜freq：definition  （括号/冒号兼容全半角，竖线只见过全角｜）
HEADER_RE = re.compile(
    r"^(?P<head>.+?)\s*[（(](?P<pos>[^（）()]+)[)）]\s*｜\s*"
    r"(?P<freq>[A-D](?:/[A-D])?)\s*[：:]\s*(?P<def>.*)$",
    re.S,
)
# 无词性词条头：headword｜freq：definition（如 "refer to｜A：…"、"prevail upon/on｜C：…"）
# ｜ 在原文中只用于词条头，普通搭配片段不会包含，可安全匹配
NOPOS_HEADER_RE = re.compile(
    r"^(?P<head>[A-Za-z][^｜（）()]*?)\s*｜\s*"
    r"(?P<freq>[A-D](?:/[A-D])?)\s*[：:]\s*(?P<def>.*)$",
    re.S,
)
# C 格式词条头（20260716 书）：word (pos)：freq，definition —— 冒号在频度前、无｜
HEADER_C_RE = re.compile(
    r"^(?P<head>.+?)\s*[（(](?P<pos>[^（）()]+)[)）]\s*[：:]\s*"
    r"(?P<freq>[A-D](?:/[A-D])?)\s*[，,]\s*(?P<def>.*)$",
    re.S,
)
NOPOS_HEADER_C_RE = re.compile(
    r"^(?P<head>[A-Za-z][^｜（）()]*?)\s*[：:]\s*"
    r"(?P<freq>[A-D](?:/[A-D])?)\s*[，,]\s*(?P<def>.*)$",
    re.S,
)
ALL_HEADER_RES = (HEADER_RE, NOPOS_HEADER_RE, HEADER_C_RE, NOPOS_HEADER_C_RE)
# C 格式探测（用于 is_blob_entry）
FORMAT_C_DETECT_RE = re.compile(r"[（(][^（）()]+[)）]\s*[：:]\s*[A-D](/[A-D])?\s*[，,]")
# 段内【标记】预切分（lookahead，不丢字符）
INLINE_MARKER_SPLIT_RE = re.compile(
    r"(?=【(?:使用提醒|区别|补充说明|词族补充|搭配|限制|共同核心义|纠错)】)"
)
# 【词族补充】内部的小词条头（冒号与释义可选，如 "persuasion (n.)｜B；persuasive (adj.)｜A/B。"）
FAMILY_HEADER_RE = re.compile(
    r"(?P<head>[^；;]+?)\s*[（(](?P<pos>[^（）()]+)[)）]\s*｜\s*"
    r"(?P<freq>[A-D](?:/[A-D])?)\s*[：:]?"
)
FAMILY_SEP_CHARS = set("；;、 \t/")

MARKERS = {
    "【使用提醒】": "usage_note",
    "【区别】": "difference",
    "【补充说明】": "group_note",
    "【词族补充】": "family_additions",
    "【共同核心义】": "group_note",
    "【纠错】": "group_note_keep_label",
    "【搭配】": "collocation",
    "【限制】": "usage_note",
}
# 不带【】的裸标记（必须带冒号才匹配，防止误伤正文）
BARE_MARKERS = {
    "区别": "difference",
    "补充说明": "group_note",
    "共同义": "group_note_keep_label",
    "共同主题": "group_note_keep_label",
}


def strip_marker(segment: str, marker: str) -> tuple[str, bool]:
    """剥掉标记与至多一个前导冒号，返回 (内容, 是否无损)。"""
    rest = segment[len(marker):]
    if rest.startswith(("：", ":")):
        content = rest[1:]
    else:
        content = rest
    rebuilt = marker + (rest[:1] if rest.startswith(("：", ":")) else "") + content
    return content, rebuilt == segment


def match_bare_marker(segment: str) -> tuple[str, str] | None:
    """匹配裸标记（区别：/ 共同义： 等），返回 (标记原文, 类别) 或 None。"""
    for bare, kind in BARE_MARKERS.items():
        for colon in ("：", ":"):
            prefix = bare + colon
            if segment.startswith(prefix):
                return prefix, kind
    return None


def parse_family_additions(content: str) -> tuple[list[dict], bool]:
    """解析【词族补充】内容为 family_additions 列表。

    无损校验：从原文中移除所有词条头与释义切片后，剩余字符必须全是分隔符。
    无释义的小词条（"persuasion (n.)｜B；"）释义归一化为空串，尾部句号随分隔符剥掉。
    """
    matches = list(FAMILY_HEADER_RE.finditer(content))
    if not matches:
        return [], False
    items: list[dict] = []
    spans: list[tuple[int, int]] = []
    strip_chars = "".join(FAMILY_SEP_CHARS) + "。"
    for i, m in enumerate(matches):
        def_start = m.end()
        def_end = matches[i + 1].start() if i + 1 < len(matches) else len(content)
        definition = content[def_start:def_end].strip(strip_chars)
        items.append(
            {
                "headword": m.group("head").strip(),
                "pos": m.group("pos").strip(),
                "frequency": m.group("freq").strip(),
                "definition": definition,
            }
        )
        spans.append((m.start(), m.end()))
        spans.append((def_start, def_end))
    residue = list(content)
    for start, end in spans:
        for j in range(start, end):
            residue[j] = ""
    leftover = "".join(residue)
    lossless = all(ch in FAMILY_SEP_CHARS or ch == "。" for ch in leftover)
    # 词条头里的 headword 若本身含分隔符开头（前一个释义残留的；），strip 掉
    for item in items:
        item["headword"] = item["headword"].strip("".join(FAMILY_SEP_CHARS))
    return items, lossless


def is_blob_entry(entry: dict) -> bool:
    items = entry.get("items")
    if not isinstance(items, list) or len(items) != 1:
        return False
    definition = (items[0] or {}).get("definition") or ""
    if "｜" in definition or FORMAT_C_DETECT_RE.search(definition):
        return True
    return any(marker in definition for marker in MARKERS)


def split_segments(blob: str) -> list[str] | None:
    """先按 " / " 切，再按段内【标记】预切；任一步丢字则返回 None。"""
    rough = blob.split(SEGMENT_SEP)
    if SEGMENT_SEP.join(rough) != blob:
        return None
    pieces: list[str] = []
    for seg in rough:
        parts = [p for p in INLINE_MARKER_SPLIT_RE.split(seg) if p]
        if "".join(parts) != seg:
            return None
        pieces.extend(parts)
    return pieces


def match_header(seg: str):
    for regex in ALL_HEADER_RES:
        m = regex.match(seg)
        if m:
            return m
    return None


def restructure_entry(entry: dict) -> tuple[dict | None, str]:
    """把 blob 词条重组成结构化词条。返回 (新词条或 None, 说明)。"""
    blob = entry["items"][0]["definition"]
    segments = split_segments(blob)
    if segments is None:
        return None, "split-join 校验失败"

    items: list[dict] = []
    difference = entry.get("difference") or ""
    group_note = entry.get("group_note") or ""
    family_additions: list[dict] = list(entry.get("family_additions") or [])

    def apply_group_note(label: str, content: str) -> None:
        nonlocal group_note
        text = f"{label}{content}" if label else content
        group_note = (group_note + " " + text).strip() if group_note else text

    i = 0
    pending_labels: list[str] = []

    def take_label_prefix() -> str:
        """取出待挂的小节标签（拼到下一个词条的释义开头，保持原位可见）。"""
        nonlocal pending_labels
        if not pending_labels:
            return ""
        prefix = "".join(pending_labels)
        pending_labels = []
        return prefix

    while i < len(segments):
        seg = segments[i]

        # 1) 带【】标记
        marker_hit = next((m for m in MARKERS if seg.startswith(m)), None)
        if marker_hit:
            content, lossless = strip_marker(seg, marker_hit)
            if not lossless:
                return None, f"标记剥离校验失败: {seg[:30]}"
            kind = MARKERS[marker_hit]
            if kind == "usage_note":
                if not items:
                    return None, "【使用提醒】出现在任何词条头之前"
                prev = items[-1].get("usage_note") or ""
                items[-1]["usage_note"] = (prev + " " + content).strip() if prev else content
            elif kind == "difference":
                difference = (difference + " " + content).strip() if difference else content
            elif kind == "group_note":
                group_note = (group_note + " " + content).strip() if group_note else content
            elif kind == "group_note_keep_label":
                had_colon = seg[len(marker_hit):].startswith(("：", ":"))
                apply_group_note(marker_hit, ("：" + content) if had_colon else content)
            elif kind == "collocation":
                if not items:
                    return None, "【搭配】出现在任何词条头之前"
                items[-1]["collocations"].append(content)
            else:  # family_additions
                fam_items, fam_lossless = parse_family_additions(content)
                if not fam_lossless:
                    return None, f"【词族补充】解析有损: {content[:40]}"
                if not fam_items:
                    return None, f"【词族补充】无词条头: {content[:40]}"
                family_additions.extend(fam_items)
            i += 1
            continue

        # 2) 裸标记（区别：/ 共同义： 等）
        bare_hit = match_bare_marker(seg)
        if bare_hit:
            prefix, kind = bare_hit
            content = seg[len(prefix):]
            if kind == "difference":
                difference = (difference + " " + content).strip() if difference else content
            elif kind == "group_note":
                group_note = (group_note + " " + content).strip() if group_note else content
            else:  # keep_label：保留「共同义：」字样进 group_note
                apply_group_note(prefix, content)
            i += 1
            continue

        # 3) 词条头（｜格式 / C 格式，带词性 / 无词性共四种形态）
        header = match_header(seg)
        if header:
            items.append(
                {
                    "headword": header.group("head").strip(),
                    "pos": (header.groupdict().get("pos") or "").strip(),
                    "frequency": header.group("freq").strip(),
                    "definition": (take_label_prefix() + header.group("def").strip()),
                    "collocations": [],
                    "usage_note": "",
                }
            )
            i += 1
            continue

        # 3b) 通用小节标签（【核心词族：vary】【相关动词】【疲惫与耗尽】等）：
        #     不是内容而是结构标记，暂存后拼到下一个词条的释义开头
        if re.match(r"^【[^】]+】", seg):
            pending_labels.append(seg)
            i += 1
            continue

        # 4) 词头里含 " / " 被切开的特殊情况（如 "skepticism / scepticism (n.)｜B：…"）：
        #    当前片段是裸词、且下一个片段是词条头 → 合并为词头前缀
        if not items:
            if i + 1 < len(segments):
                nxt = segments[i + 1]
                header2 = match_header(nxt)
                if header2 and "【" not in seg:
                    items.append(
                        {
                            "headword": (seg + SEGMENT_SEP + header2.group("head").strip()).strip(),
                            "pos": (header2.groupdict().get("pos") or "").strip(),
                            "frequency": header2.group("freq").strip(),
                            "definition": (take_label_prefix() + header2.group("def").strip()),
                            "collocations": [],
                            "usage_note": "",
                        }
                    )
                    i += 2
                    continue

            # 4b) 词条本身已有结构化词头（如 doom：items[0] 已有 headword/pos/freq），
            #     blob 只是释义里粘了标记 → 以原词头为容器，首段作释义
            orig = entry["items"][0]
            if orig.get("headword"):
                items.append(
                    {
                        "headword": orig.get("headword") or "",
                        "pos": orig.get("pos") or "",
                        "frequency": orig.get("frequency") or "",
                        "definition": (take_label_prefix() + seg),
                        "collocations": [],
                        "usage_note": "",
                    }
                )
                i += 1
                continue
            return None, f"首个片段不是词条头且无原词头: {seg[:40]}"

        # 5) 普通片段 → 当前词的例句/搭配
        items[-1]["collocations"].append(take_label_prefix() + seg)
        i += 1

    if pending_labels:
        # blob 以标签结尾的罕见情况：不丢，挂到 group_note
        leftover = "".join(pending_labels)
        group_note = (group_note + " " + leftover).strip() if group_note else leftover

    if not items:
        return None, "未解析出任何词条头"

    new_entry = dict(entry)
    new_entry["items"] = items
    new_entry["difference"] = difference
    new_entry["group_note"] = group_note
    new_entry["family_additions"] = family_additions
    return new_entry, "ok"


def compute_content_sha256(doc: dict) -> str:
    """与 wordlist_engine.compute_content_sha256 完全一致的哈希口径。"""
    payload = {
        "subtitle": doc.get("subtitle") if isinstance(doc.get("subtitle"), str) else "",
        "corrections": doc.get("corrections") if isinstance(doc.get("corrections"), list) else [],
        "sections": doc.get("sections") if isinstance(doc.get("sections"), list) else [],
    }
    canonical = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


def count_entries_words(doc: dict) -> tuple[int, int]:
    entry_count = 0
    word_count = 0
    for section in doc.get("sections") or []:
        entries = section.get("entries") or []
        entry_count += len(entries)
        for entry in entries:
            word_count += len(entry.get("items") or [])
    return entry_count, word_count


def main() -> int:
    write_mode = "--write" in sys.argv
    report: list[str] = []
    report.append("# blob 词条重组报告")
    report.append("")
    report.append(f"模式：{'**写盘**' if write_mode else 'dry-run（未写盘）'}")
    report.append("")

    total_fixed = 0
    total_skipped = 0
    total_words_before = 0
    total_words_after = 0
    changed_books: list[dict] = []

    for book_dir in sorted(HOMEPAGE_WORDBOOKS.iterdir()):
        if not book_dir.is_dir():
            continue
        wf = book_dir / "wordbook.json"
        if not wf.exists():
            continue
        doc = json.loads(wf.read_text(encoding="utf-8"))
        fixed = 0
        skipped: list[str] = []
        words_before = 0
        words_after = 0

        for section in doc.get("sections") or []:
            for idx, entry in enumerate(section.get("entries") or []):
                words_before += len(entry.get("items") or [])
                if not is_blob_entry(entry):
                    words_after += len(entry.get("items") or [])
                    continue
                new_entry, reason = restructure_entry(entry)
                if new_entry is None:
                    skipped.append(f"{entry.get('group', '?')}: {reason}")
                    words_after += len(entry.get("items") or [])
                    continue
                entry.clear()
                entry.update(new_entry)
                fixed += 1
                words_after += len(entry.get("items") or [])

        if fixed == 0 and not skipped:
            continue

        total_fixed += fixed
        total_skipped += len(skipped)
        total_words_before += words_before
        total_words_after += words_after
        report.append(f"## {book_dir.name}")
        report.append(f"- 重组成功：{fixed} 条；跳过：{len(skipped)} 条")
        report.append(f"- 逐词数：{words_before} → {words_after}")
        for line in skipped:
            report.append(f"  - ⚠️ {line}")
        report.append("")

        if write_mode and fixed > 0:
            doc["content_sha256"] = compute_content_sha256(doc)
            doc["entry_count"], doc["word_count"] = count_entries_words(doc)
            wf.write_text(json.dumps(doc, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            changed_books.append(
                {
                    "archive_id": doc["archive_id"],
                    "content_sha256": doc["content_sha256"],
                    "entry_count": doc["entry_count"],
                    "word_count": doc["word_count"],
                }
            )

    report.append("## 总计")
    report.append(f"- 重组成功：{total_fixed} 条；跳过：{total_skipped} 条")
    report.append(f"- 全库逐词数：{total_words_before} → {total_words_after}")

    if write_mode and changed_books:
        manifest_path = HOMEPAGE_WORDBOOKS / "wordbooks-manifest.json"
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        by_id = {b["archive_id"]: b for b in changed_books}
        updated = 0
        for row in manifest.get("wordbooks") or []:
            patch = by_id.get(row.get("archive_id"))
            if patch:
                row["content_sha256"] = patch["content_sha256"]
                row["entry_count"] = patch["entry_count"]
                row["word_count"] = patch["word_count"]
                updated += 1
        manifest_path.write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
        report.append(f"- manifest 更新：{updated} 行")

    REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
    REPORT_PATH.write_text("\n".join(report) + "\n", encoding="utf-8")
    print(f"fixed={total_fixed} skipped={total_skipped} words {total_words_before}->{total_words_after}")
    print(f"report: {REPORT_PATH}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
