from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import List, Tuple
import difflib
import re
import sys
from zipfile import ZipFile

from lxml import etree
from openpyxl import Workbook

# 两份文档需为「同一 Word 文档的不同版本」：相似度低于此阈值则视为两份无关文档，不生成 Excel。
# 可根据实际文档长度与改版幅度微调（建议范围 0.30～0.45）。
MIN_PAIR_SIMILARITY = 0.36


class ComparisonRejectedError(Exception):
    """用于表示两份文档相似度过低，拒绝进行差异导出。"""


@dataclass
class DiffRecord:
    category: str  # 段落 / 表格
    location: str
    change_type: str  # 新增 / 删除 / 修改
    old_text: str
    new_text: str


def normalize_text(text: str) -> str:
    """归一化文本，减少格式差异导致的误报。"""
    text = text.replace("\u3000", " ")
    text = re.sub(r"\s+", " ", text).strip()
    # 忽略中文字符之间的空格差异：例如 "其 他" -> "其他"
    text = re.sub(r"(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])", "", text)
    # 忽略数字前后空格：例如 "第 1 条" -> "第1条"、"100 元" -> "100元"
    text = re.sub(r"(?<=\d)\s+(?=\D)", "", text)
    text = re.sub(r"(?<=\D)\s+(?=\d)", "", text)
    # 忽略斜杠前后空格：例如 "A / B" -> "A/B"
    text = re.sub(r"\s*/\s*", "/", text)
    # 忽略标点前多余空格：例如 "其他： 。" -> "其他：。"
    text = re.sub(r"\s+([,.;:!?，。；：！？、）】》」』])", r"\1", text)
    # 忽略左侧标点后的多余空格：例如 "（ 内容" -> "（内容"
    text = re.sub(r"([（【《「『])\s+", r"\1", text)
    return text


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def read_document_root(doc_path: Path) -> etree._Element:
    """直接读取 document.xml，避免漏掉内容控件中的文本。"""
    with ZipFile(doc_path, "r") as zf:
        xml_bytes = zf.read("word/document.xml")
    return etree.fromstring(xml_bytes)


def get_node_text(node: etree._Element) -> str:
    """拼接一个节点下的全部文本片段（w:t）。"""
    texts = node.xpath(".//w:t/text()", namespaces=NS)
    return normalize_text("".join(texts))


def _walk_blocks(
    node: etree._Element,
    paragraphs: List[str],
    cells: List[Tuple[str, str]],
    table_counter: List[int],
) -> None:
    """
    深度遍历正文块级节点：
    - 识别段落 w:p
    - 识别表格 w:tbl
    - 递归处理内容控件 w:sdt 中的内容
    """
    for child in node:
        tag = etree.QName(child).localname

        if tag == "p":
            paragraph_text = get_node_text(child)
            if paragraph_text:
                paragraphs.append(paragraph_text)
            continue

        if tag == "tbl":
            table_counter[0] += 1
            table_idx = table_counter[0]

            rows = child.xpath("./w:tr", namespaces=NS)
            for r_idx, row in enumerate(rows, start=1):
                table_cells = row.xpath("./w:tc", namespaces=NS)
                for c_idx, cell in enumerate(table_cells, start=1):
                    cell_text = get_node_text(cell)
                    location = f"表格{table_idx}-行{r_idx}-列{c_idx}"
                    cells.append((location, cell_text))
            continue

        # 内容控件中通常包裹正文内容（w:sdtContent）
        if tag == "sdt":
            content_nodes = child.xpath("./w:sdtContent", namespaces=NS)
            for content in content_nodes:
                _walk_blocks(content, paragraphs, cells, table_counter)
            continue

        # 常见容器节点继续下钻，确保不漏文本
        if tag in {"body", "tc", "tr", "sdtContent"}:
            _walk_blocks(child, paragraphs, cells, table_counter)


def extract_paragraphs_and_cells(doc_path: Path) -> Tuple[List[str], List[Tuple[str, str]]]:
    paragraphs: List[str] = []
    cells: List[Tuple[str, str]] = []
    table_counter = [0]

    root = read_document_root(doc_path)
    body_nodes = root.xpath("./w:body", namespaces=NS)
    if body_nodes:
        _walk_blocks(body_nodes[0], paragraphs, cells, table_counter)

    return paragraphs, cells


def pair_similarity(
    old_paragraphs: List[str],
    old_cells: List[Tuple[str, str]],
    new_paragraphs: List[str],
    new_cells: List[Tuple[str, str]],
) -> float:
    """
    估算两份文档是否为同一 Word 文档的改版：综合「段落顺序」与「全文拼接」相似度，取较高值。
    完全无关的两份文档两者通常都很低。
    """
    if not old_paragraphs and not new_paragraphs:
        return 0.0

    para_ratio = 0.0
    if old_paragraphs and new_paragraphs:
        para_ratio = difflib.SequenceMatcher(
            a=old_paragraphs, b=new_paragraphs, autojunk=False
        ).ratio()

    old_cell_texts = [t for _, t in old_cells if t]
    new_cell_texts = [t for _, t in new_cells if t]
    old_full = "\n".join(old_paragraphs + old_cell_texts)
    new_full = "\n".join(new_paragraphs + new_cell_texts)
    char_ratio = 0.0
    if old_full and new_full:
        char_ratio = difflib.SequenceMatcher(
            None, old_full, new_full, autojunk=False
        ).ratio()

    return max(para_ratio, char_ratio)


def compare_sequences(
    old_items: List[str],
    new_items: List[str],
    category: str,
    location_prefix: str,
) -> List[DiffRecord]:
    """
    比较两个一维序列（例如段落列表），输出增删改记录。
    基于 difflib 的 opcode 拆分，支持 replace / insert / delete。
    """
    records: List[DiffRecord] = []
    matcher = difflib.SequenceMatcher(a=old_items, b=new_items, autojunk=False)

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue

        old_chunk = old_items[i1:i2]
        new_chunk = new_items[j1:j2]
        max_len = max(len(old_chunk), len(new_chunk))

        for offset in range(max_len):
            old_text = old_chunk[offset] if offset < len(old_chunk) else ""
            new_text = new_chunk[offset] if offset < len(new_chunk) else ""

            if old_text and new_text:
                change_type = "修改"
            elif old_text and not new_text:
                change_type = "删除"
            else:
                change_type = "新增"

            location = f"{location_prefix}{max(i1, j1) + offset + 1}"
            records.append(
                DiffRecord(
                    category=category,
                    location=location,
                    change_type=change_type,
                    old_text=old_text,
                    new_text=new_text,
                )
            )
    return records


def compare_table_cells(
    old_cells: List[Tuple[str, str]],
    new_cells: List[Tuple[str, str]],
) -> List[DiffRecord]:
    """
    按单元格顺序比较表格内容。
    单元格位置相同但文本不同 => 修改；
    旧有新无 => 删除；旧无新有 => 新增。
    """
    records: List[DiffRecord] = []
    max_len = max(len(old_cells), len(new_cells))

    for idx in range(max_len):
        old_loc, old_text = old_cells[idx] if idx < len(old_cells) else ("", "")
        new_loc, new_text = new_cells[idx] if idx < len(new_cells) else ("", "")

        # 仅空格差异在提取阶段已归一化；这里相同文本（含都为空）直接忽略。
        if old_text == new_text:
            continue

        if idx < len(old_cells) and idx < len(new_cells):
            records.append(
                DiffRecord(
                    category="表格",
                    location=new_loc or old_loc,
                    change_type="修改",
                    old_text=old_text,
                    new_text=new_text,
                )
            )
        elif idx < len(old_cells):
            records.append(
                DiffRecord(
                    category="表格",
                    location=old_loc,
                    change_type="删除",
                    old_text=old_text,
                    new_text="",
                )
            )
        else:
            records.append(
                DiffRecord(
                    category="表格",
                    location=new_loc,
                    change_type="新增",
                    old_text="",
                    new_text=new_text,
                )
            )

    return records


def write_to_excel(records: List[DiffRecord], output_path: Path) -> None:
    wb = Workbook()

    # 仅保留“修改项”sheet（按需求去掉“新增项”“删除项”）。
    ws_modify = wb.active
    ws_modify.title = "修改项"

    headers = ["类别", "位置", "变更类型", "原文", "新文"]
    ws_modify.append(headers)

    for record in records:
        row = [
            record.category,
            record.location,
            record.change_type,
            record.old_text,
            record.new_text,
        ]
        ws_modify.append(row)

    wb.save(str(output_path))


def compare_contract_files(old_doc: Path, new_doc: Path, output_excel: Path) -> int:
    """
    比较两份 Word 文档并导出 Excel。
    返回值为差异条目数量；当判定不是同一文档改版时抛出 ComparisonRejectedError。
    """
    old_paragraphs, old_cells = extract_paragraphs_and_cells(old_doc)
    new_paragraphs, new_cells = extract_paragraphs_and_cells(new_doc)

    sim = pair_similarity(old_paragraphs, old_cells, new_paragraphs, new_cells)
    if sim < MIN_PAIR_SIMILARITY:
        raise ComparisonRejectedError(
            "拒绝生成差异结果：两份文档整体相似度过低，疑似并非同一文档的不同版本。\n"
            f"当前相似度约为 {sim:.2%}，阈值 {MIN_PAIR_SIMILARITY:.0%}。\n"
            "请确认放对了「第一版 / 第二版」文件，或调整脚本中的 MIN_PAIR_SIMILARITY。"
        )

    paragraph_diffs = compare_sequences(
        old_paragraphs, new_paragraphs, category="段落", location_prefix="段落"
    )
    table_diffs = compare_table_cells(old_cells, new_cells)
    all_diffs = paragraph_diffs + table_diffs
    write_to_excel(all_diffs, output_excel)
    return len(all_diffs)


def get_runtime_dir() -> Path:
    """
    脚本直接运行时：以 .py 所在目录为基准。
    PyInstaller 单文件 exe：以 exe 所在目录为基准（文档与结果与 exe 同目录）。
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def detect_two_contract_docs(base_dir: Path) -> Tuple[Path, Path]:
    """
    自动识别当前目录下的两份 Word 文档 docx（排除 Word 临时文件 ~$*）。
    识别顺序采用文件名升序：第一个=第一版，第二个=第二版。
    """
    candidates = sorted(
        p
        for p in base_dir.glob("*.docx")
        if p.is_file() and not p.name.startswith("~$")
    )

    if len(candidates) != 2:
        names = "\n".join(f"- {p.name}" for p in candidates) or "- （未检测到）"
        raise FileNotFoundError(
            "当前目录应当且仅应包含 2 个 Word 文档 docx 文件。\n"
            f"实际检测到 {len(candidates)} 个：\n{names}\n"
            "请保留两份文档后重试。"
        )

    def _name_score(name: str, keywords: List[str]) -> int:
        return sum(1 for k in keywords if k in name)

    first_keywords = ["第一", "基准", "原版", "初版", "限制", "v1", "版本1", "old"]
    second_keywords = ["第二", "修改", "改编", "新版", "v2", "版本2", "new"]

    a, b = candidates[0], candidates[1]
    a_first = _name_score(a.name.lower(), first_keywords)
    a_second = _name_score(a.name.lower(), second_keywords)
    b_first = _name_score(b.name.lower(), first_keywords)
    b_second = _name_score(b.name.lower(), second_keywords)

    # 优先依据文件名语义判断“第一版/第二版”，避免纯字典序导致方向反转。
    if a_first > a_second and b_second > b_first:
        return a, b
    if b_first > b_second and a_second > a_first:
        return b, a

    # 无法判断时使用文件名字典序，保证行为稳定可预期。
    return a, b


def main() -> None:
    base_dir = get_runtime_dir()
    old_doc, new_doc = detect_two_contract_docs(base_dir)
    output_excel = base_dir / "文档差异结果.xlsx"
    print(f"第一份文档（基准版）：{old_doc.name}")
    print(f"第二份文档（修改版）：{new_doc.name}")
    try:
        diff_count = compare_contract_files(old_doc, new_doc, output_excel)
    except ComparisonRejectedError as exc:
        print(str(exc))
        sys.exit(1)

    print(f"比对完成，共发现 {diff_count} 处差异。")
    print(f"结果文件：{output_excel}")


if __name__ == "__main__":
    main()
