"""
스킬(.claude/skills/hwpx-report-filler)의 실행 진입점.
raw/ 에서 템플릿·엑셀·내용파일을 고정 파일명 없이 자동으로 찾아 out/ 에 보고서를 생성한다.

raw/ 에 있어야 하는 것:
  - 서식이 있는 HWPX 템플릿 1개 ("문서 표준 양식.hwpx" 구조, 제목 자리에 "보고자료 제목"이 그대로 있는 파일)
  - 서식은 무시하고 내용만 쓴 HWPX 1개 (템플릿 문구를 앵커로 써서 작성)
  - 표에 들어갈 엑셀(.xlsx) 1개
"""
from pathlib import Path

from hwpx_fill import (
    build_output_filename,
    extract_table,
    generate_hwpx,
    parse_template_style_content,
    read_xlsx_table,
)

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "raw"
OUT = ROOT / "out"

TEMPLATE_MARKER = "보고자료 제목"


def _read_section0_text(hwpx_path: Path) -> str:
    import zipfile

    with zipfile.ZipFile(hwpx_path) as z:
        return z.read("Contents/section0.xml").decode("utf-8")


def find_inputs(need_content: bool = True, need_excel: bool = True) -> tuple[Path, Path | None, Path | None]:
    """raw/ 에서 템플릿·내용파일·엑셀을 찾는다. 모자라거나 겹치면 바로 에러를 낸다.
    need_content/need_excel 이 False 면 그 파일은 찾지 않는다(Claude가 정리한 JSON을 쓰거나, 표 없이 만들 때)."""
    hwpx_files = sorted(RAW.glob("*.hwpx"))
    xlsx_files = sorted(RAW.glob("*.xlsx"))

    if need_excel and len(xlsx_files) != 1:
        raise FileNotFoundError(
            f"raw/ 에 엑셀(.xlsx) 파일이 정확히 1개 있어야 합니다 (현재 {len(xlsx_files)}개)."
        )

    template_path = None
    content_path = None
    for p in hwpx_files:
        if TEMPLATE_MARKER in _read_section0_text(p):
            if template_path is not None:
                raise FileNotFoundError(f"템플릿으로 보이는 .hwpx가 2개 이상입니다: {template_path.name}, {p.name}")
            template_path = p
        elif need_content:
            if content_path is not None:
                raise FileNotFoundError(f"내용파일로 보이는 .hwpx가 2개 이상입니다: {content_path.name}, {p.name}")
            content_path = p

    if template_path is None:
        raise FileNotFoundError("raw/ 에 서식 템플릿(.hwpx)이 없습니다.")
    if need_content and content_path is None:
        raise FileNotFoundError(
            "raw/ 에 내용만 담은 .hwpx 가 없습니다. 맨 앞 두 줄은 제목·날짜, 그 뒤로 "
            "'ㅇ (배경)' 'ㅇ (목적)' 'ㅇ (내용)' '표 제목' '참고1' 문구 다음 줄에 내용을 적어주세요."
        )
    return template_path, content_path, (xlsx_files[0] if need_excel else None)


def main(content_override: dict | None = None) -> Path:
    # JSON에 "include_table": false 를 주면 엑셀 없이 표 제목·표를 뺀 보고서를 만든다.
    include_table = (content_override or {}).get("include_table", True)
    template_path, content_path, excel_path = find_inputs(
        need_content=content_override is None, need_excel=include_table
    )

    # 내용 파일이 "ㅇ (배경)" 같은 템플릿 문구를 그대로 따라 쓴 경우엔 자동으로 읽어 쓰고,
    # 자유롭게 쓴 경우엔 Claude가 직접 읽고 판단한 content dict(JSON)를 받아 쓴다.
    content = content_override if content_override is not None else parse_template_style_content(content_path)
    if include_table:
        header, data_rows = extract_table(read_xlsx_table(excel_path))
        content = dict(content, table_header=header, table_rows=data_rows)

    out_path = OUT / build_output_filename(content["title"])
    generate_hwpx(template_path, out_path, content)
    print(f"생성 완료: {out_path}")
    return out_path


if __name__ == "__main__":
    import json
    import sys

    override = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8")) if len(sys.argv) > 1 else None
    main(override)
