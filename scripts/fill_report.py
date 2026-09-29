"""
PPT Report Data Auto-Filler

데이터/*.txt 파일을 읽어 보고서/*.pptx 템플릿(플레이스홀더가 아직 채워지지 않은 원본 서식)의
표와 차트에 값을 채워 넣는다. PPT 변환은 하지 않고, 같은 파일 경로에 덮어쓴다.

사용법:
    python3 scripts/fill_report.py
    (데이터 폴더에 존재하는 전체 텍스트 파일을 대상으로 처리)
"""
from __future__ import annotations

import copy
import glob
import logging
import os
import re

from pptx import Presentation
from pptx.chart.data import CategoryChartData

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger("fill_report")

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(REPO_ROOT, "데이터")
REPORT_DIR = os.path.join(REPO_ROOT, "보고서")
REPORT_NAME_TMPL = "[디와이덕양]분기보고서_{ym}.pptx"

# 데이터 파일 내 각 블록(순서 고정)의 컬럼 정의 - 헤더 행이 없는 경우의 폴백으로 사용
BLOCK_HEADERS = [
    [
        "PRODUCT_COUNT", "COO_Y_COUNT", "COO_N_COUNT", "COO_ERR_COUNT",
        "COO_YT_RATIO", "COO_Y_RATIO", "COO_N_RATIO", "COO_ERROR_RATIO",
        "CTC_NO_COUNT", "RVC_NO_COUNT", "COMBINE_NO_COUNT", "COO_N_COUNT2",
        "CTC_NO_RATIO", "RVC_NO_RATIO", "COMBINE_NO_RATIO", "COO_NT_RATIO",
    ],
    ["ERROR_MSG", "CNT"],
    ["FTA_NAME", "PRODUCT_COUNT", "COO_Y_COUNT", "COO_N_COUNT", "COO_ERR_COUNT", "COO_Y_RATIO"],
    ["VENDOR_NAME", "ITEM_COUNT", "COO_T_COUNT", "COO_Y_COUNT", "COO_N_COUNT", "NOT_COUNT", "COO_RATIO", "PO_AMOUNT"],
    ["VENDOR_NAME", "ITEM_CODE", "HS_CODE", "UNIT_PRICE", "PO_AMOUNT", "COO_N"],
]

ERROR_CATEGORIES = ["BOM 누락", "HS코드 누락", "원재료 단가", "기타"]


def _looks_like_header(fields: list[str], expected: list[str]) -> bool:
    non_empty = [f for f in fields if f.strip()]
    if not non_empty:
        return False
    matches = sum(1 for f in non_empty if f.strip() in expected)
    return matches >= max(1, len(non_empty) // 2)


def split_blocks(text: str) -> list[list[str]]:
    blocks: list[list[str]] = []
    current: list[str] | None = None
    for line in text.splitlines():
        if line.startswith("<") and ">" in line:
            if current is not None:
                blocks.append(current)
            current = []
        elif current is not None:
            current.append(line)
    if current is not None:
        blocks.append(current)
    return blocks


def parse_block_rows(lines: list[str], expected_header: list[str]) -> list[dict]:
    rows = []
    for line in lines:
        if not line.strip():
            continue
        fields = line.split("\t")
        if _looks_like_header(fields, expected_header):
            continue
        # 첫 필드는 순번(인덱스) 컬럼 -> 제거
        values = [f.strip() for f in fields[1:]]
        if len(values) < len(expected_header):
            log.warning("데이터 컬럼 수 부족 (기대 %d, 실제 %d): %r", len(expected_header), len(values), line)
            continue
        row = dict(zip(expected_header, values))
        rows.append(row)
    return rows


def parse_data_file(path: str) -> dict:
    text = open(path, encoding="utf-8").read()
    blocks = split_blocks(text)
    if len(blocks) < 5:
        raise ValueError(f"예상된 5개 블록을 찾지 못했습니다 (발견: {len(blocks)}개): {path}")

    summary_rows = parse_block_rows(blocks[0], BLOCK_HEADERS[0])
    if not summary_rows:
        raise ValueError(f"'원산지 판정 및 비역내산 판정 상세' 데이터가 없습니다: {path}")
    summary = summary_rows[0]

    errors = parse_block_rows(blocks[1], BLOCK_HEADERS[1])
    fta = parse_block_rows(blocks[2], BLOCK_HEADERS[2])
    vendors = parse_block_rows(blocks[3], BLOCK_HEADERS[3])
    materials = parse_block_rows(blocks[4], BLOCK_HEADERS[4])

    return {
        "summary": summary,
        "errors": errors,
        "fta": fta,
        "vendors": vendors,
        "materials": materials,
    }


def _to_float(value: str, default: float = 0.0) -> float:
    try:
        return float(value.replace(",", ""))
    except (ValueError, AttributeError):
        return default


def set_cell_text(cell, text: str) -> None:
    """기존 Run의 서식을 유지하면서 셀 텍스트만 교체한다."""
    tf = cell.text_frame
    p0 = tf.paragraphs[0]
    runs = p0.runs
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        p0.text = text
    for extra_p in tf.paragraphs[1:]:
        extra_p._p.getparent().remove(extra_p._p)


def get_tables(slide):
    return [s.table for s in slide.shapes if s.has_table]


def get_charts(slide):
    return [s.chart for s in slide.shapes if s.has_chart]


def ensure_data_rows(table, needed: int) -> None:
    """표의 데이터 행(헤더 제외) 수를 needed 개로 맞춘다 (마지막 행을 복제/삭제)."""
    current = len(table.rows) - 1
    if current == needed:
        return
    tbl_elem = table._tbl
    if current < needed:
        last_tr = list(table.rows)[-1]._tr
        for _ in range(needed - current):
            tbl_elem.append(copy.deepcopy(last_tr))
    else:
        for _ in range(current - needed):
            last_tr = list(table.rows)[-1]._tr
            tbl_elem.remove(last_tr)


def map_error_categories(errors: list[dict]) -> dict[str, float]:
    totals = {c: 0.0 for c in ERROR_CATEGORIES}
    for row in errors:
        msg = row.get("ERROR_MSG", "").strip()
        cnt = _to_float(row.get("CNT", "0"))
        if msg == "BOM 누락":
            totals["BOM 누락"] += cnt
        elif msg == "HS코드 누락":
            totals["HS코드 누락"] += cnt
        elif msg == "소요량 또는 금액이 0 인 것이 존재합니다.":
            totals["원재료 단가"] += cnt
        else:
            totals["기타"] += cnt
    return totals


def fill_report(template_path: str, data: dict, output_path: str) -> None:
    prs = Presentation(template_path)
    slides = list(prs.slides)

    # ---- Slide 3: 원산지 판정 및 비역내산 판정 상세 ----
    try:
        slide3 = slides[2]
        tables3 = get_tables(slide3)
        table1, table2 = tables3[0], tables3[1]
        s = data["summary"]

        table1_rows = [
            ("전체 품목 수", s["PRODUCT_COUNT"], s["COO_YT_RATIO"]),
            ("역내산 판정 건수", s["COO_Y_COUNT"], s["COO_Y_RATIO"]),
            ("비역내산 판정 건수", s["COO_N_COUNT"], s["COO_N_RATIO"]),
            ("판정 실패 건수", s["COO_ERR_COUNT"], s["COO_ERROR_RATIO"]),
        ]
        for i, (name, cnt, ratio) in enumerate(table1_rows, start=1):
            row = table1.rows[i]
            set_cell_text(row.cells[0], name)
            set_cell_text(row.cells[1], cnt)
            set_cell_text(row.cells[2], ratio)

        table2_rows = [
            ("세번변경기준 미충족", s["CTC_NO_COUNT"], s["CTC_NO_RATIO"]),
            ("부가가치기준 미충족", s["RVC_NO_COUNT"], s["RVC_NO_RATIO"]),
            ("조합기준 미충족", s["COMBINE_NO_COUNT"], s["COMBINE_NO_RATIO"]),
            ("비역내산 합계", s["COO_N_COUNT2"], s["COO_NT_RATIO"]),
        ]
        for i, (name, cnt, ratio) in enumerate(table2_rows, start=1):
            row = table2.rows[i]
            set_cell_text(row.cells[0], name)
            set_cell_text(row.cells[1], cnt)
            set_cell_text(row.cells[2], ratio)

        charts3 = get_charts(slide3)
        chart1, chart2 = charts3[0], charts3[1]

        cd1 = CategoryChartData()
        cd1.categories = ["역내산율", "역외산율", "판정실패"]
        cd1.add_series("", (
            _to_float(s["COO_Y_RATIO"]) / 100,
            _to_float(s["COO_N_RATIO"]) / 100,
            _to_float(s["COO_ERROR_RATIO"]) / 100,
        ))
        chart1.replace_data(cd1)

        cd2 = CategoryChartData()
        cd2.categories = ["세번변경기준", "부가가치기준", "조합기준"]
        cd2.add_series("", (
            _to_float(s["CTC_NO_RATIO"]) / 100,
            _to_float(s["RVC_NO_RATIO"]) / 100,
            _to_float(s["COMBINE_NO_RATIO"]) / 100,
        ))
        chart2.replace_data(cd2)
    except Exception:
        log.exception("슬라이드 3(원산지 판정 및 비역내산 판정 상세) 처리 중 오류 발생: %s", template_path)

    # ---- Slide 4: 원산지 판정 에러 현황 ----
    try:
        slide4 = slides[3]
        charts4 = get_charts(slide4)
        chart3 = charts4[0]
        totals = map_error_categories(data["errors"])
        cd3 = CategoryChartData()
        cd3.categories = ERROR_CATEGORIES
        cd3.add_series("특혜", tuple(totals[c] for c in ERROR_CATEGORIES))
        chart3.replace_data(cd3)
    except Exception:
        log.exception("슬라이드 4(원산지 판정 에러 현황) 처리 중 오류 발생: %s", template_path)

    # ---- Slide 5 & 6: 협정별 원산지 판정 현황 ----
    try:
        slide5 = slides[4]
        slide6 = slides[5]
        table5 = get_tables(slide5)[0]
        table6_raw = get_tables(slide6)[0]

        fta_rows = data["fta"]
        slide5_capacity = len(table5.rows) - 1
        first_part = fta_rows[:slide5_capacity]
        remaining_part = fta_rows[slide5_capacity:]

        for i, row in enumerate(first_part, start=1):
            trow = table5.rows[i]
            set_cell_text(trow.cells[0], row["FTA_NAME"])
            set_cell_text(trow.cells[1], row["PRODUCT_COUNT"])
            set_cell_text(trow.cells[2], row["COO_Y_COUNT"])
            set_cell_text(trow.cells[3], row["COO_N_COUNT"])
            set_cell_text(trow.cells[4], row["COO_ERR_COUNT"])
            set_cell_text(trow.cells[5], row["COO_Y_RATIO"])

        ensure_data_rows(table6_raw, len(remaining_part))
        table6 = get_tables(slide6)[0]
        for i, row in enumerate(remaining_part, start=1):
            trow = table6.rows[i]
            set_cell_text(trow.cells[0], row["FTA_NAME"])
            set_cell_text(trow.cells[1], row["PRODUCT_COUNT"])
            set_cell_text(trow.cells[2], row["COO_Y_COUNT"])
            set_cell_text(trow.cells[3], row["COO_N_COUNT"])
            set_cell_text(trow.cells[4], row["COO_ERR_COUNT"])
            set_cell_text(trow.cells[5], row["COO_Y_RATIO"])
    except Exception:
        log.exception("슬라이드 5/6(협정별 원산지 판정 현황) 처리 중 오류 발생: %s", template_path)

    # ---- Slide 7: 협력사별 원산지확인서 수취 현황 ----
    try:
        slide7 = slides[6]
        table7_raw = get_tables(slide7)[0]
        vendors = data["vendors"]
        ensure_data_rows(table7_raw, len(vendors))
        table7 = get_tables(slide7)[0]
        for i, row in enumerate(vendors, start=1):
            trow = table7.rows[i]
            set_cell_text(trow.cells[0], row["VENDOR_NAME"])
            set_cell_text(trow.cells[1], row["ITEM_COUNT"])
            set_cell_text(trow.cells[2], row["COO_T_COUNT"])
            set_cell_text(trow.cells[3], row["COO_Y_COUNT"])
            set_cell_text(trow.cells[4], row["COO_N_COUNT"])
            set_cell_text(trow.cells[5], row["NOT_COUNT"])
            set_cell_text(trow.cells[6], row["COO_RATIO"])
            set_cell_text(trow.cells[7], row["PO_AMOUNT"])
    except Exception:
        log.exception("슬라이드 7(협력사별 원산지확인서 수취 현황) 처리 중 오류 발생: %s", template_path)

    # ---- Slide 8: 협력사별 원산지확인서 수취/미수취 자재 내역 ----
    try:
        slide8 = slides[7]
        table8_raw = get_tables(slide8)[0]
        materials = data["materials"]
        ensure_data_rows(table8_raw, len(materials))
        table8 = get_tables(slide8)[0]
        for i, row in enumerate(materials, start=1):
            trow = table8.rows[i]
            set_cell_text(trow.cells[0], row["VENDOR_NAME"])
            set_cell_text(trow.cells[1], row["ITEM_CODE"])
            set_cell_text(trow.cells[2], row["HS_CODE"])
            set_cell_text(trow.cells[3], row["UNIT_PRICE"])
            set_cell_text(trow.cells[4], row["PO_AMOUNT"])
            set_cell_text(trow.cells[5], row["COO_N"])
    except Exception:
        log.exception("슬라이드 8(협력사별 원산지확인서 수취/미수취 자재 내역) 처리 중 오류 발생: %s", template_path)

    prs.save(output_path)
    log.info("저장 완료: %s", output_path)


def main() -> None:
    data_files = sorted(glob.glob(os.path.join(DATA_DIR, "*.txt")))
    if not data_files:
        log.warning("데이터 폴더에 텍스트 파일이 없습니다: %s", DATA_DIR)
        return

    for data_path in data_files:
        ym = os.path.splitext(os.path.basename(data_path))[0]
        if not re.fullmatch(r"\d{6}", ym):
            log.warning("파일명이 YYYYMM 형식이 아니어서 건너뜁니다: %s", data_path)
            continue

        report_path = os.path.join(REPORT_DIR, REPORT_NAME_TMPL.format(ym=ym))
        if not os.path.exists(report_path):
            log.error("대응하는 보고서 템플릿을 찾지 못했습니다: %s", report_path)
            continue

        try:
            data = parse_data_file(data_path)
        except Exception:
            log.exception("데이터 파일 파싱 실패: %s", data_path)
            continue

        try:
            fill_report(report_path, data, report_path)
        except Exception:
            log.exception("보고서 채우기 실패: %s -> %s", data_path, report_path)


if __name__ == "__main__":
    main()
