"""
raw/ 의 HWPX 템플릿 + 엑셀 표를 읽어, 서식을 그대로 둔 채 내용만 채운
새 HWPX 파일을 out/ 에 생성한다. 외부 라이브러리 설치 없이 표준 라이브러리만 사용한다.

템플릿(문서 표준 양식.hwpx)의 서식(글꼴·크기 등)은 기존 문단을 그대로 복제해 재사용하고,
개수가 정해지지 않은 부분은 필요한 만큼 늘리거나 줄인다 (02_기획안.md 처리 1~5단계):
  - 제목 칸, 날짜·기관 칸(없으면 오늘 날짜)
  - "ㅇ (배경)"/"ㅇ (목적)" 한 줄 + 필요한 만큼의 "-" 부연설명 줄
  - "□ 주요 내용" 아래, 라벨이 자유로운 "ㅇ (라벨) 내용" 줄 1개 이상
  - 표 제목, 표 본문(1열은 순번 자동, 그 뒤는 엑셀의 '순번' 다음 열 전부) — 엑셀의 열 개수만큼 표 열이,
    데이터 행 수만큼 표 행이 늘거나 줄어듦(표 전체 너비는 유지한 채 열 너비만 다시 균등하게 나눔)
  - 표의 모든 셀은 "한 줄로 입력"(SQUEEZE)으로 고정하고, 데이터가 ONE_PAGE_MAX_DATA_ROWS(15)행을
    넘어 2페이지 이상 이어지면 표의 "글자처럼 취급"을 해제해 고정 위치에 띄움(1페이지 안에서 끝나면 유지)
  - "참고1" 아래 참고자료 제목 칸 (참고 내용이 없으면 이 문단 전체를 삭제)
"""
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import date
from pathlib import Path

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
HP_NS = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"

# 템플릿 자체에 이미 있는 문구를 "자리표시"로 인식한다. 이 문구를 그대로 써주면
# 바로 다음 줄을 해당 칸의 내용으로 가져온다 (공백을 지우고 비교해 "ㅇ(배경)"도 인정).
ANCHOR_TO_KEY = {
    "ㅇ(배경)": "background_detail",
    "ㅇ(목적)": "purpose_detail",
    "ㅇ(내용)": "content_summary",
    "표제목": "table_caption",
    "참고1": "ref_title",
}


def _esc(text: str) -> str:
    return (text or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def read_xlsx_table(xlsx_path: Path):
    """엑셀 첫 시트를 [[셀,...], ...] 형태로 읽는다 (shared strings/inline string 모두 지원)."""
    with zipfile.ZipFile(xlsx_path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall(f"{NS}si"):
                shared.append("".join(t.text or "" for t in si.iter(f"{NS}t")))

        root = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
        rows = []
        for row in root.iter(f"{NS}row"):
            cells = {}
            for c in row.findall(f"{NS}c"):
                ref = c.get("r") or ""
                col_letters = "".join(ch for ch in ref if ch.isalpha())
                col_idx = 0
                for ch in col_letters:
                    col_idx = col_idx * 26 + (ord(ch.upper()) - 64)
                col_idx -= 1

                t = c.get("t")
                v_el = c.find(f"{NS}v")
                is_el = c.find(f"{NS}is")
                if t == "s" and v_el is not None:
                    val = shared[int(v_el.text)]
                elif t == "inlineStr" and is_el is not None:
                    val = "".join(tt.text or "" for tt in is_el.iter(f"{NS}t"))
                elif v_el is not None:
                    val = v_el.text or ""
                else:
                    val = ""
                cells[col_idx] = val
            if cells:
                maxcol = max(cells)
                rows.append([cells.get(i, "") for i in range(maxcol + 1)])
        return rows


def extract_plain_paragraphs(hwpx_path: Path) -> list[str]:
    """서식은 무시하고, HWPX 최상위 문단의 텍스트만 순서대로 뽑아낸다 (표/도형 등 중첩 문단은 제외)."""
    with zipfile.ZipFile(hwpx_path) as z:
        root = ET.fromstring(z.read("Contents/section0.xml"))
    paragraphs = []
    for p in root.findall(f"{HP_NS}p"):
        texts = [t.text or "" for t in p.iter(f"{HP_NS}t")]
        line = "".join(texts).strip()
        if line:
            paragraphs.append(line)
    return paragraphs


def parse_template_style_content(hwpx_path: Path) -> dict:
    """새 라벨을 외우지 않고, 템플릿에 이미 있던 문구(ㅇ (배경) 등)를 그대로 써서
    구분한 내용 파일을 읽는다. 맨 앞 두 줄은 제목·날짜로 보고, 그 뒤로는
    템플릿 문구가 나오면 다음 줄을 그 칸의 내용으로 가져온다.

    이 함수는 "ㅇ (배경)" 문구를 정확히 그대로 쓴, 가장 단순한 경우만 처리한다.
    실제로는 사람마다 "ㅇ 배경 : 내용"처럼 자유롭게 쓰는 경우가 많으므로, 이 함수가
    실패하면(ValueError) Claude가 extract_plain_paragraphs()로 직접 읽고 판단해
    아래 fill_section0() 의 content 형식(JSON)으로 만들어 쓴다."""
    lines = extract_plain_paragraphs(hwpx_path)
    found = {}
    if len(lines) >= 1:
        found["title"] = lines[0]
    if len(lines) >= 2:
        found["date_org"] = lines[1]

    i = 2
    while i < len(lines):
        normalized = "".join(lines[i].split()).strip("[]")
        key = ANCHOR_TO_KEY.get(normalized)
        if key and i + 1 < len(lines):
            found[key] = lines[i + 1].strip()
            i += 2
        else:
            i += 1

    required = ["title", "background_detail", "purpose_detail", "content_summary", "table_caption"]
    missing = [k for k in required if k not in found]
    if missing:
        raise ValueError(
            f"내용 파일에서 다음을 찾지 못함: {missing}. "
            "맨 앞 두 줄은 제목·날짜, 그 뒤로는 'ㅇ (배경)' 'ㅇ (목적)' 'ㅇ (내용)' '표 제목' '참고1' "
            "문구 다음 줄에 내용을 적어주세요. (날짜·참고1 은 없어도 됨)"
        )
    return {
        "title": found["title"],
        "date_org": found.get("date_org"),
        "overview_items": [
            {"depth": 1, "label": "배경", "text": found["background_detail"]},
            {"depth": 1, "label": "목적", "text": found["purpose_detail"]},
        ],
        "main_items": [{"depth": 1, "label": "내용", "text": found["content_summary"]}],
        "table_caption": found["table_caption"],
        "ref_title": found.get("ref_title", ""),
    }


KOREAN_WEEKDAYS = ["월", "화", "수", "목", "금", "토", "일"]


def format_date_org(org: str = "대외협력실", when: date | None = None) -> str:
    """'yy.m.d.(요일), 기관' 형식의 날짜·기관 줄을 오늘 날짜로 만든다."""
    when = when or date.today()
    weekday = KOREAN_WEEKDAYS[when.weekday()]
    return f"{when.strftime('%y')}.{when.month}.{when.day}.({weekday}), {org}"


def _top_level_paragraph_spans(xml_text: str) -> list:
    """최상위 문단(hp:p)들의 (시작,끝) 위치를 모두 찾는다. 표/도형 안에 또 있는
    hp:p는 깊이를 세어 걸러낸다."""
    depth = 0
    start = None
    spans = []
    for m in re.finditer(r"<hp:p(?=[ >])|</hp:p>", xml_text):
        if m.group(0) == "</hp:p>":
            depth -= 1
            if depth == 0 and start is not None:
                spans.append((start, m.end()))
                start = None
        else:
            if depth == 0:
                start = m.start()
            depth += 1
    return spans


def _find_paragraph_span(xml_text: str, contains: str) -> tuple:
    for start, end in _top_level_paragraph_spans(xml_text):
        if contains in xml_text[start:end]:
            return start, end
    raise ValueError(f"문단을 찾지 못함: {contains!r}")


def remove_top_level_paragraph(xml_text: str, contains: str) -> str:
    """최상위 문단(hp:p) 중 contains 문자열을 포함하는 문단을 통째로 지운다."""
    start, end = _find_paragraph_span(xml_text, contains)
    return xml_text[:start] + xml_text[end:]


def _set_first_t_text(block: str, new_text: str) -> str:
    """문단/셀 블록 안의 첫 <hp:t>...</hp:t> 내용만 바꾼다 (서식·구조는 그대로)."""
    i = block.index("<hp:t>") + len("<hp:t>")
    j = block.index("</hp:t>", i)
    return block[:i] + _esc(new_text) + block[j:]


def _set_last_t_text(block: str, new_text: str) -> str:
    """문단 블록 안의 마지막 <hp:t>...</hp:t> 내용만 바꾼다 (여러 run으로 나뉜 문단용)."""
    j = block.rindex("</hp:t>")
    i = block.rindex("<hp:t>", 0, j) + len("<hp:t>")
    return block[:i] + _esc(new_text) + block[j:]


def _bullet_paragraph(stamp: str, label: str, text: str) -> str:
    """"ㅇ (라벨) 내용" 한 줄(깊이 1)을 stamp(기존 ㅇ 문단)에서 서식을 그대로 복제해 만든다."""
    new_text = f" ㅇ ({label}) {text}".rstrip() if text else f" ㅇ ({label})"
    return _set_first_t_text(stamp, new_text)


def _dash_paragraph(stamp: str, text: str) -> str:
    """"  - 내용" 부연설명 한 줄(깊이 2)을 stamp(기존 "-" 문단)에서 복제해 만든다."""
    return _set_first_t_text(stamp, f"  - {text}")


STAR_TEXT_CHARPR = "31"  # 13pt 맑은고딕 (날짜·표 데이터 칸과 같은 스타일)


def _star_paragraph(stamp: str, text: str) -> str:
    """"   * 내용" 한 줄(깊이 3)을 stamp(기존 "*" 문단)에서 복제해 만든다.
    "*" 기호는 별도 run이라 그대로 두고, 기호 뒤 내용 run만 13pt 맑은고딕으로 바꿔 글을 채운다."""
    block = stamp.replace('charPrIDRef="33"', f'charPrIDRef="{STAR_TEXT_CHARPR}"', 1)
    return _set_last_t_text(block, f" {text}" if text else " ")


def render_outline(items: list, stamps: dict, spacer: str) -> str:
    """ㅁ(섹션) 아래의 ㅇ/-/* 줄들을 순서대로 그려낸다. items 는 문서에 나오는 순서 그대로
    [{"depth": 1|2|3, "label": str(깊이1=ㅇ만 씀), "text": str}, ...] 형태의 "평평한" 목록이다.
    depth 1=ㅇ, 2=-, 3=*. 중간 단계를 생략해도(예: ㅇ 다음에 바로 depth 3) 그대로 순서대로 넣으면 된다."""
    builders = {
        1: lambda item: _bullet_paragraph(stamps["bullet"], item.get("label", ""), item.get("text", "")),
        2: lambda item: _dash_paragraph(stamps["dash"], item.get("text", "")),
        3: lambda item: _star_paragraph(stamps["star"], item.get("text", "")),
    }
    parts = []
    for i, item in enumerate(items):
        if i > 0:
            parts.append(spacer)
        parts.append(builders[item.get("depth", 1)](item))
    return "".join(parts)


def extract_table(rows: list) -> tuple:
    """엑셀에서 '순번' 열 다음에 오는 모든 열을 그대로 표 데이터로 가져온다(개수를 4개로 고정하지 않음).
    '순번'이 적힌 행을 헤더로 보고(없으면 첫 행), 그 열의 위치를 기준으로 뒤에 오는 열을 전부 쓴다.
    앞에 안내용 제목 행이나 순번 앞의 빈 열이 있어도 상관없다. 완전히 빈 데이터 행은 건너뛴다."""
    header_idx = 0
    for i, row in enumerate(rows):
        if "순번" in row:
            header_idx = i
            break
    header_row = rows[header_idx]
    seq_idx = header_row.index("순번")
    header = list(header_row[seq_idx + 1 :])
    while header and header[-1] == "":
        header.pop()
    n_cols = len(header)
    data_rows = []
    for row in rows[header_idx + 1 :]:
        slice_ = row[seq_idx + 1 : seq_idx + 1 + n_cols]
        if any(cell.strip() for cell in slice_):
            data_rows.append(slice_)
    return header, data_rows


MAIN_TABLE_ID = "1849124803"
SEQ_COL_WIDTH = 4265  # 템플릿 1열(순번) 칸 너비 — 열 개수를 바꿔도 그대로 유지
TOTAL_DATA_WIDTH = 42228  # 템플릿 표 전체 너비(46493) - 순번 칸(4265). 데이터 열들이 나눠 쓰는 폭
LINESEG_WIDTH_MARGIN = 1021  # 템플릿에서 확인한 cellSz width와 lineseg horzsize의 고정 차이(10557-9536, 4265-3244)


def _set_cell_width(tc_xml: str, width: int) -> str:
    """셀 하나의 cellSz width와 안쪽 글자 줄바꿈 폭(horzsize)을 같은 너비로 맞춘다."""
    tc_xml = re.sub(r'(<hp:cellSz width=")\d+(")', rf"\g<1>{width}\g<2>", tc_xml, count=1)
    horzsize = max(width - LINESEG_WIDTH_MARGIN, 1)
    tc_xml = re.sub(r'(horzsize=")\d+(")', rf"\g<1>{horzsize}\g<2>", tc_xml, count=1)
    return tc_xml


def adjust_table_cols(xml_text: str, target_data_cols: int) -> str:
    """메인 표(1열 순번 + 데이터 열)의 데이터 열 개수를 target_data_cols 개로 맞춘다.
    표 전체 너비(46493)는 그대로 두고, 데이터 열 너비(42228)만 열 개수에 맞춰 다시 균등하게 나눈다.
    모자라면 마지막 데이터 열을 복제해 늘리고, 많으면 뒤에서부터 지운다."""
    m = re.search(rf'<hp:tbl id="{MAIN_TABLE_ID}".*?</hp:tbl>', xml_text, re.DOTALL)
    if m is None:
        raise ValueError(f"표(id={MAIN_TABLE_ID})를 찾지 못함")
    tbl_block = m.group(0)

    trs = [mm.group(0) for mm in re.finditer(r"<hp:tr>.*?</hp:tr>", tbl_block, re.DOTALL)]
    tcs_per_row = [re.findall(r"<hp:tc[ >].*?</hp:tc>", tr, re.DOTALL) for tr in trs]
    n_existing = len(tcs_per_row[0]) - 1  # 0번째 칸은 순번

    base = TOTAL_DATA_WIDTH // target_data_cols
    remainder = TOTAL_DATA_WIDTH - base * target_data_cols
    widths = [base] * target_data_cols
    widths[-1] += remainder

    new_trs = []
    for tr, tcs in zip(trs, tcs_per_row):
        seq_tc, data_tcs = tcs[0], list(tcs[1:])
        if target_data_cols <= n_existing:
            data_tcs = data_tcs[:target_data_cols]
        else:
            last = data_tcs[-1]
            for col in range(n_existing, target_data_cols):
                clone = re.sub(r'colAddr="\d+"', f'colAddr="{col + 1}"', last, count=1)
                data_tcs.append(clone)
        data_tcs = [_set_cell_width(tc, widths[i]) for i, tc in enumerate(data_tcs)]
        new_trs.append(tr.replace("".join(tcs), seq_tc + "".join(data_tcs), 1))

    new_tbl_block = re.sub(r'colCnt="\d+"', f'colCnt="{target_data_cols + 1}"', tbl_block, count=1)
    new_tbl_block = new_tbl_block.replace("".join(trs), "".join(new_trs), 1)
    return xml_text[: m.start()] + new_tbl_block + xml_text[m.end() :]


ONE_PAGE_MAX_DATA_ROWS = 15  # 이 템플릿 기준 실측값: 데이터 15행까지는 1페이지, 16행부터 2페이지로 넘어감


def _apply_table_squeeze(xml_text: str) -> str:
    """메인 표의 모든 셀을 "한 줄로 입력"(글자 간격을 줄여 한 줄에 맞춤)으로 바꾼다."""
    m = re.search(rf'<hp:tbl id="{MAIN_TABLE_ID}".*?</hp:tbl>', xml_text, re.DOTALL)
    if m is None:
        raise ValueError(f"표(id={MAIN_TABLE_ID})를 찾지 못함")
    tbl_block = m.group(0).replace('lineWrap="BREAK"', 'lineWrap="SQUEEZE"')
    return xml_text[: m.start()] + tbl_block + xml_text[m.end() :]


def _set_table_treat_as_char(xml_text: str, treat_as_char: bool) -> str:
    """표가 ONE_PAGE_MAX_DATA_ROWS를 넘어 2페이지 이상으로 이어지면 "글자처럼 취급"을 해제해
    고정된 위치에 띄운다(위치는 지금 자리 그대로 — vertRelTo/horzRelTo/offset은 그대로 두고
    treatAsChar만 바꾼다). 1페이지 안에서 끝나면 "글자처럼 취급"을 그대로 유지한다."""
    m = re.search(rf'<hp:tbl id="{MAIN_TABLE_ID}".*?</hp:tbl>', xml_text, re.DOTALL)
    if m is None:
        raise ValueError(f"표(id={MAIN_TABLE_ID})를 찾지 못함")
    tbl_block = m.group(0)
    value = "1" if treat_as_char else "0"
    new_tbl_block = re.sub(r'(<hp:pos treatAsChar=")\d+(")', rf"\g<1>{value}\g<2>", tbl_block, count=1)
    return xml_text[: m.start()] + new_tbl_block + xml_text[m.end() :]


def _set_row_number_cell(data_row_xml: str, number: int) -> str:
    """표 데이터 행의 1열(순번) 칸 숫자를 바꾼다."""
    m = re.search(r"<hp:tc[ >].*?</hp:tc>", data_row_xml, re.DOTALL)
    first_cell = m.group(0)
    new_cell = re.sub(r"(<hp:t>)\d+(</hp:t>)", rf"\g<1>{number}\g<2>", first_cell, count=1)
    return data_row_xml[: m.start()] + new_cell + data_row_xml[m.end() :]


def adjust_table_rows(xml_text: str, target_data_rows: int) -> str:
    """메인 표(5열, 1열 순번)의 데이터 행 수를 target_data_rows 개로 맞춘다.
    모자라면 마지막 행을 복제해 번호를 이어 붙이고, 많으면 뒤에서부터 지운다."""
    m = re.search(rf'<hp:tbl id="{MAIN_TABLE_ID}".*?</hp:tbl>', xml_text, re.DOTALL)
    if m is None:
        raise ValueError(f"표(id={MAIN_TABLE_ID})를 찾지 못함")
    tbl_block = m.group(0)

    trs = [mm.group(0) for mm in re.finditer(r"<hp:tr>.*?</hp:tr>", tbl_block, re.DOTALL)]
    header_tr, data_trs = trs[0], trs[1:]
    n_existing = len(data_trs)

    if target_data_rows <= n_existing:
        new_data_trs = data_trs[:target_data_rows]
    else:
        new_data_trs = list(data_trs)
        last = data_trs[-1]
        for row_num in range(n_existing + 1, target_data_rows + 1):
            clone = last.replace(f'rowAddr="{n_existing}"', f'rowAddr="{row_num}"')
            clone = _set_row_number_cell(clone, row_num)
            new_data_trs.append(clone)

    new_tbl_block = re.sub(r'rowCnt="\d+"', f'rowCnt="{target_data_rows + 1}"', tbl_block, count=1)
    new_tbl_block = new_tbl_block.replace("".join(trs), header_tr + "".join(new_data_trs), 1)
    return xml_text[: m.start()] + new_tbl_block + xml_text[m.end() :]


def fill_section0(xml_text: str, content: dict) -> str:
    """템플릿 section0.xml 텍스트에 내용을 채워 새 XML 문자열을 반환한다.

    content 형식:
      title, date_org(없으면 오늘 날짜), table_caption, ref_title(없으면 "참고1" 문단 삭제)
      overview_items: "□ 추진 개요" 아래에 들어갈 ㅇ/-/* 줄 목록 (문서에 나오는 순서 그대로)
      main_items:     "□ 주요 내용" 아래에 들어갈 ㅇ/-/* 줄 목록 (라벨 자유, 1개 이상)
      각 줄은 {"depth": 1|2|3, "label": str(깊이1=ㅇ일 때만), "text": str} 형태.
      depth 1=ㅇ, 2=-, 3=*. 중간 단계를 생략해도(예: ㅇ 다음에 바로 depth 3) 순서대로 넣으면 된다.
      table_header, table_rows: 표에 채울 데이터 (열 개수·행 수 모두 데이터에 맞춰 표가 늘어나거나 줄어듦)
    """

    ref_title = (content.get("ref_title") or "").strip()
    if not ref_title:
        xml_text = remove_top_level_paragraph(xml_text, "참고1")

    # "ㅇ (배경)"/"-"/"*" 문단 자체를 각 깊이의 stamp로, 그 바로 다음 빈 문단을 줄 사이 여백으로 쓴다.
    bg_start, bg_end = _find_paragraph_span(xml_text, " ㅇ (배경)")
    bullet_stamp = xml_text[bg_start:bg_end]
    spacer = xml_text[xml_text.index("<hp:p", bg_end) : xml_text.index("</hp:p>", bg_end) + len("</hp:p>")]
    dash_start, dash_end = _find_paragraph_span(xml_text, "  - ")
    dash_stamp = xml_text[dash_start:dash_end]
    star_start, star_end = _find_paragraph_span(xml_text, 'charPrIDRef="40"')
    star_stamp = xml_text[star_start:star_end]
    stamps = {"bullet": bullet_stamp, "dash": dash_stamp, "star": star_stamp}

    overview_items = content.get("overview_items") or []
    purpose_span = _find_paragraph_span(xml_text, " ㅇ (목적) ")
    new_region = render_outline(overview_items, stamps, spacer)
    xml_text = xml_text[:bg_start] + new_region + xml_text[purpose_span[1]:]

    # "□ 주요 내용" 아래 "ㅇ (내용)" 한 줄을, main_items 개요(ㅇ/-/* 트리)로 바꾼다.
    main_items = content.get("main_items") or [{"label": "내용", "text": content.get("content_summary", "")}]
    content_start, content_end = _find_paragraph_span(xml_text, " ㅇ (내용) ")
    new_main = render_outline(main_items, stamps, spacer)
    xml_text = xml_text[:content_start] + new_main + xml_text[content_end:]

    replacements = [
        ("보고자료 제목", _esc(content["title"])),
        ("‘26.10.7.(수), KIAT 대외협력실", _esc(content.get("date_org") or format_date_org())),
        ("[표 제목]", _esc(content["table_caption"])),
    ]
    if ref_title:
        replacements.append((" 참고자료 제목", " " + _esc(ref_title)))

    for old, new in replacements:
        count = xml_text.count(old)
        if count != 1:
            raise ValueError(f"placeholder not found exactly once: {old!r} (count={count})")
        xml_text = xml_text.replace(old, new, 1)

    header = content.get("table_header", [])
    data_rows = content.get("table_rows", [])
    n_cols = len(header)
    xml_text = adjust_table_cols(xml_text, n_cols)
    xml_text = adjust_table_rows(xml_text, len(data_rows))
    xml_text = _apply_table_squeeze(xml_text)
    xml_text = _set_table_treat_as_char(xml_text, treat_as_char=len(data_rows) <= ONE_PAGE_MAX_DATA_ROWS)

    def fill_cell(match: "re.Match", value: str) -> str:
        block = match.group(0)
        run = re.search(r'<hp:run charPrIDRef="(\d+)"/>', block)
        if run is None or not value:
            return block
        charpr = run.group(1)
        old_run = run.group(0)
        new_run = f'<hp:run charPrIDRef="{charpr}"><hp:t>{_esc(value)}</hp:t></hp:run>'
        return block.replace(old_run, new_run, 1)

    def table_sub(match: "re.Match") -> str:
        block = match.group(0)
        addr = re.search(r'<hp:cellAddr colAddr="(\d+)" rowAddr="(\d+)"', block)
        if addr is None:
            return block
        col, row = int(addr.group(1)), int(addr.group(2))
        if col == 0 or col > n_cols:
            return block
        if row == 0:
            value = header[col - 1] if col - 1 < len(header) else ""
        elif 1 <= row <= len(data_rows):
            data_row = data_rows[row - 1] if row - 1 < len(data_rows) else []
            value = data_row[col - 1] if col - 1 < len(data_row) else ""
        else:
            return block
        return fill_cell(match, value)

    xml_text = re.sub(r"<hp:tc[ >].*?</hp:tc>", table_sub, xml_text, flags=re.DOTALL)
    return xml_text


def generate_hwpx(template_path: Path, out_path: Path, content: dict) -> None:
    with zipfile.ZipFile(template_path) as zin:
        names = zin.namelist()
        section0 = zin.read("Contents/section0.xml").decode("utf-8")
        new_section0 = fill_section0(section0, content)

        out_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in names:
                data = zin.read(name)
                if name == "Contents/section0.xml":
                    data = new_section0.encode("utf-8")
                # mimetype은 압축하지 않아야 하는 관례를 따름
                compress = zipfile.ZIP_STORED if name == "mimetype" else zipfile.ZIP_DEFLATED
                zout.writestr(name, data, compress_type=compress)


def build_output_filename(title: str, today: date | None = None) -> str:
    today = today or date.today()
    ymd = today.strftime("%y%m%d")
    safe_title = re.sub(r'[\\/:*?"<>|]', "_", title)
    return f"{ymd}_{safe_title}.hwpx"
