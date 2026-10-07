"""
raw/ 에 아직 "내용만 담은 한글파일"이 없을 때, 1차 검증을 돌려보기 위한
최소 샘플 내용 파일을 표준 라이브러리(zipfile)만으로 만든다.
실제 파일(한글에서 라벨: 내용 형식으로 작성해 raw/ 에 저장한 .hwpx)이 들어오면 이 스크립트는 필요 없다.
"""
import zipfile
from pathlib import Path

SECTION0 = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<hp:sec xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
{paragraphs}
</hp:sec>"""

PARA_TEMPLATE = '<hp:p><hp:run><hp:t>{text}</hp:t></hp:run></hp:p>'


def _esc(text: str) -> str:
    return (text or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def make_sample_content(path: Path, lines: list[str]) -> None:
    paragraphs = "\n".join(PARA_TEMPLATE.format(text=_esc(line)) for line in lines)
    section0 = SECTION0.format(paragraphs=paragraphs)
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("Contents/section0.xml", section0)


if __name__ == "__main__":
    make_sample_content(
        Path(__file__).resolve().parent.parent / "raw" / "내용_샘플.hwpx",
        lines=[
            "3주차 실습 자동화 테스트 보고",
            "'26.10.7.(수), 대외협력실",
            "ㅇ (배경)",
            "반복되는 HWPX 보고서 작성 시 서식이 깨지는 불편이 있었음",
            "ㅇ (목적)",
            "서식 유지한 채 내용만 채우는 최소 기능을 검증하기 위함",
            "ㅇ (내용)",
            "템플릿 구조를 분석해 제목·본문·표 칸에 내용을 자동으로 채움",
            "표 제목",
            "[후보기업(기관) 목록]",
            "참고1",
            "02_기획안.md, 01_페인포인트.md",
        ],
    )
    print("샘플 내용 파일 생성 완료")
