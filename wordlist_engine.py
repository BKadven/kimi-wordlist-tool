from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Iterator

from docx import Document
from docx.document import Document as DocumentObject
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.shared import Cm, Pt, RGBColor
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
from openai import OpenAI

BASE_URL = "https://api.moonshot.cn/v1"
COLUMN_WIDTHS_CM = (4.4, 2.2, 10.8)
TABLE_WIDTH_CM = sum(COLUMN_WIDTHS_CM)
SUPPORTED_INPUT_SUFFIXES = {".docx", ".txt"}
ProgressCallback = Callable[[int, str], None]


@dataclass(frozen=True)
class ProcessingResult:
    output_docx: Path
    raw_json: Path
    parsed_json: Path
    model: str
    prompt_tokens: int | None = None
    completion_tokens: int | None = None
    total_tokens: int | None = None


def emit(callback: ProgressCallback | None, percent: int, message: str) -> None:
    if callback:
        callback(percent, message)


def iter_block_items(parent: DocumentObject | _Cell) -> Iterator[Paragraph | Table]:
    """按 Word 文档中的实际顺序遍历段落和表格。"""
    if isinstance(parent, DocumentObject):
        parent_element = parent.element.body
    elif isinstance(parent, _Cell):
        parent_element = parent._tc
    else:
        raise TypeError(f"不支持的 Word 容器类型：{type(parent)!r}")

    for child in parent_element.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def extract_docx_content(path: Path) -> str:
    document = Document(path)
    blocks: list[str] = []

    for block in iter_block_items(document):
        if isinstance(block, Paragraph):
            value = block.text.strip()
            if value:
                blocks.append(value)
            continue

        for row in block.rows:
            values: list[str] = []
            for cell in row.cells:
                cell_text = "\n".join(
                    paragraph.text.strip()
                    for paragraph in cell.paragraphs
                    if paragraph.text.strip()
                )
                values.append(cell_text)

            if any(values):
                blocks.append("【原表格行】" + " ｜ ".join(values))

    content = "\n".join(blocks).strip()
    if not content:
        raise ValueError("输入 Word 文档中没有提取到可处理的文字。")
    return content


def extract_txt_content(path: Path) -> str:
    for encoding in ("utf-8-sig", "utf-8", "utf-16", "gb18030"):
        try:
            content = path.read_text(encoding=encoding).strip()
            break
        except UnicodeDecodeError:
            continue
    else:
        content = path.read_text(encoding="utf-8", errors="replace").strip()

    if not content:
        raise ValueError("输入 TXT 文档中没有提取到可处理的文字。")
    return content


def extract_input_content(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        return extract_docx_content(path)
    if suffix == ".txt":
        return extract_txt_content(path)
    raise FileNotFoundError("请选择一个有效的 .docx 或 .txt 原始单词记录文件。")


def build_messages(rules: str, source_text: str) -> list[dict[str, str]]:
    schema = r'''
{
  "document_title": "IELTS 单词本整理版",
  "subtitle": "词族成组 · 同义辨析 · 雅思实用频度逐词标注",
  "corrections": [
    {
      "original": "错误形式",
      "corrected": "正确形式",
      "note": "简短说明"
    }
  ],
  "sections": [
    {
      "title": "一、名词与名词核心词族",
      "entries": [
        {
          "group": "左栏显示的核心词条、词族或近义组",
          "frequency_range": "A-C",
          "items": [
            {
              "headword": "单词或短语",
              "pos": "n./v./adj./adv./phr. v. 等",
              "frequency": "A、B、C、D 或 A/B",
              "definition": "准确、自然的中文释义和必要语义限制",
              "collocations": ["常用搭配一", "常用搭配二"],
              "usage_note": "没有必要说明时写空字符串"
            }
          ],
          "family_additions": [
            {
              "headword": "必要的常用词族补充",
              "pos": "词性",
              "frequency": "A/B/C/D",
              "definition": "简短释义"
            }
          ],
          "difference": "同组词之间的 nuance、替换限制和语境区别",
          "group_note": "其他必要说明；没有则写空字符串"
        }
      ]
    },
    {"title": "二、动词与动词核心词族", "entries": []},
    {"title": "三、形容词与形容词核心词族", "entries": []},
    {"title": "四、副词与表达偏好", "entries": []}
  ]
}
'''.strip()

    system_message = (
        "你是一名严谨的英语词汇编辑和雅思词汇资料整理专家。"
        "必须严格执行用户提供的工作规范，并只返回合法 JSON 对象。"
        "不得返回 Markdown、代码围栏、解释性前言或结语。"
    )

    user_message = f"""
请整理下面的原始英文单词记录。

【必须执行的完整工作规范】
{rules}

【额外结构化输出要求】
1. 只能返回一个合法 JSON 对象，键名和层级严格采用下方结构。
2. sections 必须按名词、动词、形容词、副词／表达的顺序排列。
3. 词族完整性优先，不得为了分板块而拆散同一词族。
4. 原始记录中的每一个英文词、词组和固定搭配都必须至少出现在 group、items、family_additions、difference、group_note 或 corrections 之一。
5. 每个 items 项目必须单独标注雅思实用频度；frequency_range 只用于左栏概括。
6. collocations 通常每词保留 1—4 个真正常用、具有复习价值的搭配。
7. 不要为了“完整”添加大量冷僻派生词。
8. definition、usage_note、difference 和 group_note 使用简体中文。
9. corrections 只记录原始材料中真实存在的拼写、词形或记录错误；没有则返回空数组。
10. 请在输出前自行检查是否遗漏、错误合并、错误拆散或误标频度。

【必须采用的 JSON 结构】
{schema}

【原始英文单词记录开始】
{source_text}
【原始英文单词记录结束】
""".strip()

    return [
        {"role": "system", "content": system_message},
        {"role": "user", "content": user_message},
    ]


def parse_json_response(raw_text: str) -> dict[str, Any]:
    raw_text = raw_text.strip()
    try:
        data = json.loads(raw_text)
    except json.JSONDecodeError as exc:
        first = raw_text.find("{")
        last = raw_text.rfind("}")
        if first == -1 or last <= first:
            raise ValueError(f"Kimi 返回内容不是合法 JSON：{exc}") from exc
        try:
            data = json.loads(raw_text[first : last + 1])
        except json.JSONDecodeError as second_exc:
            raise ValueError(f"Kimi 返回的 JSON 无法解析：{second_exc}") from second_exc

    if not isinstance(data, dict):
        raise ValueError("Kimi 返回的 JSON 顶层不是对象。")
    return data


def validate_data(data: dict[str, Any]) -> None:
    sections = data.get("sections")
    if not isinstance(sections, list) or not sections:
        raise ValueError("整理结果缺少有效的 sections 列表。")

    entry_count = 0
    for section in sections:
        if not isinstance(section, dict):
            raise ValueError("sections 中存在非对象项目。")
        title = str(section.get("title", "")).strip()
        entries = section.get("entries", [])
        if not title:
            raise ValueError("存在没有标题的章节。")
        if not isinstance(entries, list):
            raise ValueError(f"章节“{title}”的 entries 不是列表。")

        for entry in entries:
            if not isinstance(entry, dict):
                raise ValueError(f"章节“{title}”中存在非对象词条。")
            if not str(entry.get("group", "")).strip():
                raise ValueError(f"章节“{title}”中存在缺少 group 的词条。")
            items = entry.get("items")
            if not isinstance(items, list) or not items:
                raise ValueError(f"词条“{entry.get('group')}”缺少 items。")
            entry_count += 1

    if entry_count == 0:
        raise ValueError("整理结果没有生成任何词条。")


def call_kimi(
    *,
    api_key: str,
    model: str,
    rules: str,
    source_text: str,
    thinking_enabled: bool,
) -> tuple[str, Any]:
    client = OpenAI(
        api_key=api_key,
        base_url=BASE_URL,
        timeout=900.0,
        max_retries=2,
    )

    response = client.chat.completions.create(
        model=model,
        messages=build_messages(rules, source_text),
        response_format={"type": "json_object"},
        max_completion_tokens=65536,
        extra_body={
            "thinking": {
                "type": "enabled" if thinking_enabled else "disabled"
            }
        },
    )

    if not response.choices:
        raise ValueError("Kimi 没有返回任何候选结果。")

    choice = response.choices[0]
    if choice.finish_reason == "length":
        raise ValueError("Kimi 输出达到长度上限，JSON 可能被截断。请减少输入内容后重试。")

    content = choice.message.content
    if not content:
        raise ValueError("Kimi 返回了空内容。")
    return content.strip(), response.usage


def value_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shading = tc_pr.find(qn("w:shd"))
    if shading is None:
        shading = OxmlElement("w:shd")
        tc_pr.append(shading)
    shading.set(qn("w:fill"), fill)


def set_cell_margins(cell, top: int = 90, start: int = 90, bottom: int = 90, end: int = 90) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.find(qn("w:tcMar"))
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)

    for name, value in {"top": top, "start": start, "bottom": bottom, "end": end}.items():
        node = tc_mar.find(qn(f"w:{name}"))
        if node is None:
            node = OxmlElement(f"w:{name}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_repeat_table_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    header = tr_pr.find(qn("w:tblHeader"))
    if header is None:
        header = OxmlElement("w:tblHeader")
        tr_pr.append(header)
    header.set(qn("w:val"), "true")


def set_run_font(run, *, size: float = 9, bold: bool = False, color: str | None = None) -> None:
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    run.font.bold = bold
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "宋体")
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def set_paragraph_format(paragraph, *, space_after: float = 3, line_spacing: float = 1.08) -> None:
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(space_after)
    paragraph.paragraph_format.line_spacing = line_spacing


def set_fixed_table_layout(table, total_width_cm: float) -> None:
    table.autofit = False
    tbl_pr = table._tbl.tblPr

    layout = tbl_pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")

    width = tbl_pr.find(qn("w:tblW"))
    if width is None:
        width = OxmlElement("w:tblW")
        tbl_pr.append(width)
    width.set(qn("w:w"), str(Cm(total_width_cm).twips))
    width.set(qn("w:type"), "dxa")


def set_table_grid_widths(table, widths_cm: tuple[float, ...]) -> None:
    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width_cm in widths_cm:
        column = OxmlElement("w:gridCol")
        column.set(qn("w:w"), str(Cm(width_cm).twips))
        grid.append(column)


def set_cell_width(cell, width_cm: float) -> None:
    width = Cm(width_cm)
    cell.width = width
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_width = tc_pr.find(qn("w:tcW"))
    if tc_width is None:
        tc_width = OxmlElement("w:tcW")
        tc_pr.append(tc_width)
    tc_width.set(qn("w:w"), str(width.twips))
    tc_width.set(qn("w:type"), "dxa")


def apply_column_widths(cells) -> None:
    for cell, width_cm in zip(cells, COLUMN_WIDTHS_CM):
        set_cell_width(cell, width_cm)


def format_item_detail(item: dict[str, Any]) -> str:
    parts: list[str] = []
    definition = value_text(item.get("definition"))
    if definition:
        parts.append(definition.rstrip("。"))

    collocations = item.get("collocations", [])
    if isinstance(collocations, list):
        clean = [value_text(value) for value in collocations if value_text(value)]
        if clean:
            parts.append(" / ".join(clean).rstrip("。"))

    usage_note = value_text(item.get("usage_note"))
    if usage_note:
        parts.append(f"【使用提醒】{usage_note.rstrip('。')}")

    return "。".join(parts) + ("。" if parts else "")


def add_labeled_paragraph(cell, label: str, content: str) -> None:
    if not content:
        return
    paragraph = cell.add_paragraph()
    set_paragraph_format(paragraph, space_after=4)
    label_run = paragraph.add_run(label)
    set_run_font(label_run, size=9, bold=True, color="1F4E78")
    content_run = paragraph.add_run(content)
    set_run_font(content_run, size=9)


def add_item_to_cell(cell, item: dict[str, Any], *, first: bool) -> None:
    paragraph = cell.paragraphs[0] if first else cell.add_paragraph()
    set_paragraph_format(paragraph, space_after=5)

    headword = value_text(item.get("headword")) or "未命名词条"
    pos = value_text(item.get("pos"))
    frequency = value_text(item.get("frequency"))

    label = headword
    if pos:
        label += f" ({pos})"
    if frequency:
        label += f"｜{frequency}"
    label += "："

    label_run = paragraph.add_run(label)
    set_run_font(label_run, size=9, bold=True)
    detail_run = paragraph.add_run(format_item_detail(item))
    set_run_font(detail_run, size=9)


def format_family_additions(values: Any) -> str:
    if not isinstance(values, list):
        return ""
    parts: list[str] = []
    for item in values:
        if not isinstance(item, dict):
            continue
        headword = value_text(item.get("headword"))
        if not headword:
            continue
        label = headword
        pos = value_text(item.get("pos"))
        frequency = value_text(item.get("frequency"))
        definition = value_text(item.get("definition"))
        if pos:
            label += f" ({pos})"
        if frequency:
            label += f"｜{frequency}"
        if definition:
            label += f"：{definition}"
        parts.append(label)
    return "；".join(parts)


def add_word_row(table, entry: dict[str, Any]) -> None:
    cells = table.add_row().cells
    apply_column_widths(cells)
    for cell in cells:
        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        set_cell_margins(cell)

    group_p = cells[0].paragraphs[0]
    set_paragraph_format(group_p)
    group_run = group_p.add_run(value_text(entry.get("group")))
    set_run_font(group_run, size=9, bold=True)

    set_cell_shading(cells[1], "E2F0D9")
    freq_p = cells[1].paragraphs[0]
    freq_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_format(freq_p)
    freq_run = freq_p.add_run(value_text(entry.get("frequency_range")) or "逐词")
    set_run_font(freq_run, size=9, bold=True, color="1F4E78")

    note_p = cells[1].add_paragraph()
    note_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_format(note_p)
    note_run = note_p.add_run("逐词见右栏")
    set_run_font(note_run, size=8, color="595959")

    items = entry.get("items", [])
    for index, item in enumerate(items):
        if isinstance(item, dict):
            add_item_to_cell(cells[2], item, first=index == 0)

    add_labeled_paragraph(cells[2], "【词族补充】：", format_family_additions(entry.get("family_additions")))
    add_labeled_paragraph(cells[2], "【区别】：", value_text(entry.get("difference")))
    add_labeled_paragraph(cells[2], "【补充说明】：", value_text(entry.get("group_note")))


def add_table_header(table) -> None:
    cells = table.rows[0].cells
    apply_column_widths(cells)
    set_repeat_table_header(table.rows[0])
    titles = ["词条 / 词族", "雅思频度", "中文释义、nuance 与搭配"]
    for cell, title in zip(cells, titles):
        set_cell_shading(cell, "1F4E78")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_margins(cell, top=80, bottom=80)
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_format(paragraph, space_after=0)
        run = paragraph.add_run(title)
        set_run_font(run, size=9, bold=True, color="FFFFFF")


def add_info_box(document) -> None:
    table = document.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    set_fixed_table_layout(table, TABLE_WIDTH_CM)
    set_table_grid_widths(table, (TABLE_WIDTH_CM,))

    cell = table.cell(0, 0)
    set_cell_width(cell, TABLE_WIDTH_CM)
    set_cell_shading(cell, "DDEBF7")
    set_cell_margins(cell, top=110, bottom=110, start=130, end=130)

    lines = [
        "本版标准：按雅思备考实用性逐词标注频度，不按词长或生僻程度机械分级。",
        "A｜高频必掌握：四科常见，或写作、口语改写价值高；要求会认、会搭配并能主动使用。",
        "B｜常见且值得掌握：特定话题、正式表达或学术语境有价值；熟悉主要用法并争取主动使用。",
        "C｜低频，认读为主：可能出现在阅读或窄语境中，但主动产出收益较低。",
        "D｜极低频/不建议主动使用：过时、生硬、文学化或专业范围过窄；优先采用更自然替代表达。",
        "词族原则：同一词族、近义组、反义词和形近词集中对照；关系不等同的词明确区分。",
    ]

    for index, line in enumerate(lines):
        paragraph = cell.paragraphs[0] if index == 0 else cell.add_paragraph()
        set_paragraph_format(paragraph, space_after=2)
        run = paragraph.add_run(line)
        set_run_font(run, size=8.5, bold=index == 0)


def add_corrections(document, corrections: Any) -> None:
    if not isinstance(corrections, list):
        return
    parts: list[str] = []
    for item in corrections:
        if not isinstance(item, dict):
            continue
        original = value_text(item.get("original"))
        corrected = value_text(item.get("corrected"))
        note = value_text(item.get("note"))
        if not original and not corrected:
            continue
        part = f"{original} → {corrected}"
        if note:
            part += f"（{note}）"
        parts.append(part)

    if not parts:
        return
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(6)
    paragraph.paragraph_format.space_after = Pt(6)
    title_run = paragraph.add_run("本次原始记录中的主要纠错：")
    set_run_font(title_run, size=9, bold=True, color="1F4E78")
    body_run = paragraph.add_run("；".join(parts) + "。")
    set_run_font(body_run, size=9)


def sanitize_filename(value: str) -> str:
    value = re.sub(r'[<>:"/\\|?*]+', "_", value).strip(" .")
    return value or "单词本"


def generate_word(data: dict[str, Any], source_path: Path, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    now = datetime.now()
    date_text = now.strftime("%Y-%m-%d")
    stamp = now.strftime("%Y%m%d_%H%M%S")
    stem = sanitize_filename(source_path.stem)
    output_path = output_dir / f"{stem}_雅思单词本_整理版_{stamp}.docx"

    document = Document()
    section = document.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.45)
    section.bottom_margin = Cm(1.45)
    section.left_margin = Cm(1.55)
    section.right_margin = Cm(1.55)

    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_after = Pt(2)
    title_run = title.add_run(f"IELTS 单词本 · {date_text}")
    set_run_font(title_run, size=16, bold=True, color="1F4E78")

    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.paragraph_format.space_after = Pt(8)
    subtitle_run = subtitle.add_run(
        value_text(data.get("subtitle")) or "词族成组 · 同义辨析 · 雅思实用频度逐词标注"
    )
    set_run_font(subtitle_run, size=9, color="595959")

    add_info_box(document)
    add_corrections(document, data.get("corrections"))

    for section_data in data.get("sections", []):
        if not isinstance(section_data, dict):
            continue
        entries = section_data.get("entries", [])
        if not isinstance(entries, list) or not entries:
            continue

        heading = document.add_paragraph()
        heading.paragraph_format.space_before = Pt(8)
        heading.paragraph_format.space_after = Pt(5)
        heading_run = heading.add_run(value_text(section_data.get("title")))
        set_run_font(heading_run, size=12, bold=True, color="1F4E78")

        table = document.add_table(rows=1, cols=3)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = "Table Grid"
        set_fixed_table_layout(table, TABLE_WIDTH_CM)
        set_table_grid_widths(table, COLUMN_WIDTHS_CM)
        add_table_header(table)

        for entry in entries:
            if isinstance(entry, dict):
                add_word_row(table, entry)

    footer = section.footer
    footer_p = footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_run = footer_p.add_run(f"IELTS 单词本 · {date_text}")
    set_run_font(footer_run, size=8, color="808080")

    document.save(output_path)
    return output_path


def process_wordlist(
    *,
    input_docx: Path,
    output_dir: Path,
    rules_path: Path,
    api_key: str,
    model: str,
    thinking_enabled: bool = False,
    progress_callback: ProgressCallback | None = None,
) -> ProcessingResult:
    input_docx = Path(input_docx)
    output_dir = Path(output_dir)
    rules_path = Path(rules_path)

    if input_docx.suffix.lower() not in SUPPORTED_INPUT_SUFFIXES or not input_docx.is_file():
        raise FileNotFoundError("请选择一个有效的 .docx 或 .txt 原始单词记录文件。")
    if not api_key.strip():
        raise ValueError("API Key 不能为空。")
    if not model.strip():
        raise ValueError("模型名称不能为空。")
    if not rules_path.is_file():
        raise FileNotFoundError(f"找不到规则文件：{rules_path}")

    emit(progress_callback, 8, "正在读取原始单词记录……")
    source_text = extract_input_content(input_docx)
    rules = rules_path.read_text(encoding="utf-8").strip()
    if not rules:
        raise ValueError("规则文件为空。")

    emit(progress_callback, 22, f"已提取 {len(source_text):,} 个字符，准备调用 Kimi……")
    raw_text, usage = call_kimi(
        api_key=api_key.strip(),
        model=model.strip(),
        rules=rules,
        source_text=source_text,
        thinking_enabled=thinking_enabled,
    )

    output_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    debug_stem = sanitize_filename(input_docx.stem)
    raw_path = output_dir / f"{debug_stem}_Kimi原始结果_{stamp}.json"
    raw_path.write_text(raw_text, encoding="utf-8")

    emit(progress_callback, 72, "正在校验 Kimi 返回的结构化数据……")
    data = parse_json_response(raw_text)
    validate_data(data)

    parsed_path = output_dir / f"{debug_stem}_结构化数据_{stamp}.json"
    parsed_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    emit(progress_callback, 86, "正在生成排版后的 Word 单词本……")
    output_docx = generate_word(data, input_docx, output_dir)
    emit(progress_callback, 100, "整理完成。")

    prompt_tokens = getattr(usage, "prompt_tokens", None) if usage else None
    completion_tokens = getattr(usage, "completion_tokens", None) if usage else None
    total_tokens = getattr(usage, "total_tokens", None) if usage else None

    return ProcessingResult(
        output_docx=output_docx,
        raw_json=raw_path,
        parsed_json=parsed_path,
        model=model.strip(),
        prompt_tokens=prompt_tokens,
        completion_tokens=completion_tokens,
        total_tokens=total_tokens,
    )


def list_available_models(api_key: str) -> list[str]:
    client = OpenAI(api_key=api_key.strip(), base_url=BASE_URL, timeout=60.0, max_retries=1)
    response = client.models.list()
    names = sorted({model.id for model in response.data if getattr(model, "id", None)})
    return names


def test_api(api_key: str, model: str) -> str:
    client = OpenAI(api_key=api_key.strip(), base_url=BASE_URL, timeout=60.0, max_retries=1)
    response = client.chat.completions.create(
        model=model.strip(),
        messages=[{"role": "user", "content": "只回复：连接成功"}],
        max_completion_tokens=32,
        extra_body={"thinking": {"type": "disabled"}},
    )
    content = response.choices[0].message.content if response.choices else None
    return (content or "连接成功").strip()
