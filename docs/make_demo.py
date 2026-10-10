"""README용 project_hwpx 소개 GIF 생성 (Pillow만 사용, Windows 맑은 고딕 글꼴 필요).
사용법: python docs/make_demo.py docs/demo.gif
"""
import sys
from PIL import Image, ImageDraw, ImageFont

OUT = sys.argv[1]
W, H = 960, 540
BG = (247, 248, 250)
INK = (28, 33, 40)
SUB = (96, 104, 116)
BLUE = (31, 90, 166)
LBLUE = (223, 235, 250)
GREEN = (30, 130, 76)
RED = (196, 60, 52)
LINE = (210, 214, 220)
F = "C:/Windows/Fonts/malgun.ttf"
FB = "C:/Windows/Fonts/malgunbd.ttf"
font = lambda s, b=False: ImageFont.truetype(FB if b else F, s)

frames, durs = [], []
SCENES = 7


def base(scene, title, kicker=""):
    im = Image.new("RGB", (W, H), BG)
    d = ImageDraw.Draw(im)
    d.rectangle([0, 0, W, 6], fill=BLUE)
    if kicker:
        d.text((48, 34), kicker, font=font(16, True), fill=BLUE)
    d.text((48, 56), title, font=font(30, True), fill=INK)
    # 진행 표시
    for i in range(SCENES):
        x = W // 2 - SCENES * 14 + i * 28
        d.ellipse([x, H - 26, x + 10, H - 16], fill=BLUE if i == scene else LINE)
    d.text((W - 160, H - 30), "project_hwpx", font=font(13), fill=SUB)
    return im, d


def push(im, ms):
    frames.append(im.copy())
    durs.append(ms)


def staged(build, steps, step_ms=550, hold_ms=2600):
    """build(n) 이 n단계까지 그린 이미지를 돌려주면, 0..steps 를 차례로 프레임으로 쌓는다."""
    for n in range(1, steps + 1):
        push(build(n), step_ms if n < steps else hold_ms)


def box(d, xy, fill, outline=None, r=10, w=2):
    d.rounded_rectangle(xy, radius=r, fill=fill, outline=outline, width=w)


def arrow(d, x1, y, x2, color=SUB):
    d.line([x1, y, x2 - 10, y], fill=color, width=3)
    d.polygon([(x2, y), (x2 - 12, y - 7), (x2 - 12, y + 7)], fill=color)


# ── 1. 타이틀 ─────────────────────────────────────────
def s1(n):
    im = Image.new("RGB", (W, H), BLUE)
    d = ImageDraw.Draw(im)
    d.text((70, 150), "project_hwpx", font=font(22, True), fill=(190, 212, 240))
    d.text((70, 190), "내용만 넣으면,", font=font(46, True), fill="white")
    if n >= 2:
        d.text((70, 252), "한글 보고서가 서식째 완성됩니다", font=font(46, True), fill="white")
    if n >= 3:
        d.text((70, 340), "HWPX 보고서 자동 채움 에이전트 · Claude Code 스킬", font=font(20), fill=(210, 225, 245))
        d.text((70, 372), "신흥기술영향평가세미나 실습 과제", font=font(18), fill=(190, 212, 240))
    return im


staged(s1, 3, 700, 2200)


# ── 2. 문제 ───────────────────────────────────────────
def s2(n):
    im, d = base(1, "매번 손으로 서식을 다시 맞췄습니다", "문제")
    items = [
        ("Claude 답변을 한글 양식에 붙이고", "개조식 · 글꼴(15pt 함초롱바탕 / 12pt 맑은고딕) 맞추기", "최대 30분"),
        ("엑셀 표를 붙이면 서식이 깨져서", "열 너비 · 표 크기 · 글꼴 다시 정렬", "최대 5분"),
    ]
    for i, (a, b, t) in enumerate(items[: min(n, 2)]):
        y = 130 + i * 130
        box(d, [48, y, W - 48, y + 105], "white", LINE)
        d.text((76, y + 20), a, font=font(22, True), fill=INK)
        d.text((76, y + 58), b, font=font(18), fill=SUB)
        d.text((W - 210, y + 34), t, font=font(28, True), fill=RED)
    if n >= 3:
        box(d, [48, 400, W - 48, 470], LBLUE, None)
        d.text((76, 418), "목표: 생성된 파일을 한글에서 열었을 때  수동 수정 0분", font=font(22, True), fill=BLUE)
    return im


staged(s2, 3, 900, 2800)


# ── 3. 입력 ───────────────────────────────────────────
RAW = [
    "기관장 현장소통 후보기업(기관) 조사 결과",
    "1. 개요",
    "ㅇ 배경 : 정부 방침인 현장 중심의 정책이행 성과 창출 필요",
    "* 국정과제 000, 기관장 경영철학 제0호",
    "ㅇ 목적 : 기관장 주도의 현장소통 기업(기관) 발굴",
    "2. 내용",
    "ㅇ 조사기간 : ‘26. 10. 2. (금) ~ 10. 9.(금)",
    "ㅇ 조사결과 : 4개 기업(기관) 발굴",
]
XL = [["명칭", "규모", "소재지"], ["AA전자", "대기업", "서울"], ["BB하이닉스", "대기업", "부산"], ["CC연구원", "전문연", "충남"], ["…", "…", "…"]]


def s3(n):
    im, d = base(2, "raw/ 폴더에 넣는 것 — 서식은 신경 쓰지 않아도 됩니다", "입력")
    box(d, [48, 120, 600, 490], "white", LINE)
    d.text((68, 132), "① 내용만 쓴 한글파일 (.hwpx) — 자유 형식", font=font(16, True), fill=BLUE)
    for i, line in enumerate(RAW[: min(n, len(RAW))]):
        d.text((68, 168 + i * 38), line, font=font(16), fill=INK)
    if n > len(RAW):
        box(d, [624, 120, W - 48, 340], "white", LINE)
        d.text((644, 132), "② 엑셀 표 (.xlsx)", font=font(16, True), fill=BLUE)
        for r, row in enumerate(XL):
            for c, v in enumerate(row):
                x, y = 644 + c * 88, 166 + r * 32
                d.rectangle([x, y, x + 88, y + 32], fill=(238, 240, 243) if r == 0 else "white", outline=LINE)
                d.text((x + 8, y + 6), v, font=font(14, r == 0), fill=INK)
    if n > len(RAW) + 1:
        box(d, [624, 360, W - 48, 490], "white", LINE)
        d.text((644, 372), "③ 서식 템플릿 (.hwpx)", font=font(16, True), fill=BLUE)
        d.text((644, 404), "기관 표준 보고서 양식", font=font(15), fill=SUB)
        d.text((644, 430), "글꼴·크기·표 모양의 기준", font=font(15), fill=SUB)
    return im


staged(s3, len(RAW) + 2, 380, 3200)


# ── 4. 처리 흐름 ──────────────────────────────────────
STEPS = [
    ("raw/ 입력 3종", "템플릿·내용·엑셀\n자동 식별"),
    ("Claude가 읽고 판단", "SKILL.md 절차대로\nㅁ/ㅇ/-/* 구조 정리"),
    ("hwpx_fill.py", "템플릿 문단을\n도장처럼 복제해 채움"),
    ("out/ 결과", "YYMMDD_제목.hwpx"),
]


def s4(n):
    im, d = base(3, "어떻게 동작하나", "처리")
    bw, gap, y = 190, 36, 170
    for i, (t, sub) in enumerate(STEPS[: min(n, 4)]):
        x = 48 + i * (bw + gap)
        box(d, [x, y, x + bw, y + 150], LBLUE if i in (1, 2) else "white", BLUE, 12)
        d.text((x + 16, y + 18), t, font=font(18, True), fill=BLUE)
        for k, s in enumerate(sub.split("\n")):
            d.text((x + 16, y + 62 + k * 28), s, font=font(15), fill=INK)
        if i > 0:
            arrow(d, x - gap + 4, y + 75, x - 4, BLUE)
    if n >= 5:
        d.text((48, 370), "• 새 서식을 만들지 않고 템플릿에 이미 있는 □ · ㅇ · - · * 문단을 복제 → 글꼴·들여쓰기가 그대로", font=font(16), fill=INK)
        d.text((48, 404), "• 엑셀 열 수·행 수만큼 표가 자동으로 늘고 줄어듦, 15행을 넘으면 다음 페이지로 이어지게 설정", font=font(16), fill=INK)
        d.text((48, 438), "• Python 표준 라이브러리만 사용 — 별도 설치 없음", font=font(16), fill=INK)
    return im


staged(s4, 5, 800, 4200)


# ── 5. 결과 문서 ──────────────────────────────────────
def report_page(im, d, n):
    px, py, pw, ph = 380, 92, 420, 392
    d.rectangle([px + 6, py + 6, px + pw + 6, py + ph + 6], fill=(220, 223, 228))
    d.rectangle([px, py, px + pw, py + ph], fill="white", outline=LINE)
    # 제목 상자(그라데이션)
    tx0, ty0, tx1, ty1 = px + 24, py + 22, px + pw - 24, py + 62
    for k in range(ty1 - ty0):
        t = k / (ty1 - ty0)
        c = tuple(int(a + (b - a) * abs(0.5 - t) * 2) for a, b in zip((255, 255, 255), (170, 205, 245)))
        d.line([tx0, ty0 + k, tx1, ty0 + k], fill=c)
    d.line([tx0, ty0, tx1, ty0], fill=BLUE, width=3)
    d.line([tx0, ty1, tx1, ty1], fill=BLUE, width=3)
    title = "기관장 현장소통 후보기업(기관) 조사 결과"
    f = font(15, True)
    d.text(((tx0 + tx1 - f.getlength(title)) / 2, ty0 + 9), title, font=f, fill=INK)
    d.text((px + pw - 160, py + 70), "26.10.8.(목), 대외협력실", font=font(11), fill=INK)
    lines = [
        (0, "□ 추진 개요", True), (1, "ㅇ (배경) 정부 방침인 현장 중심의 정책이행 성과 창출 필요", False),
        (2, "* 국정과제 000, 기관장 경영철학 제0호", False), (1, "ㅇ (목적) 기관장 주도의 현장소통 기업(기관) 발굴", False),
        (0, "□ 주요 내용", True), (1, "ㅇ (조사기간) '26. 10. 2.(금) ~ 10. 9.(금)", False), (1, "ㅇ (조사결과) 4개 기업(기관) 발굴", False),
    ]
    y = py + 96
    for i, (ind, s, b) in enumerate(lines[: min(n, len(lines))]):
        d.text((px + 28 + ind * 14, y), s, font=font(13 if b else 11, b), fill=INK if ind < 2 else SUB)
        y += 26 if b else 22
    if n > len(lines):
        d.text((px + pw / 2 - 60, y + 4), "[후보기업(기관) 목록]", font=font(11, True), fill=INK)
        cols = ["순번", "명칭", "규모", "종업원수", "소재지", "상세소재지", "비고"]
        rows = [cols, ["1", "AA전자", "대기업", "12480", "서울", "강남구", "기반구축"], ["2", "BB하이닉스", "대기업", "8930", "부산", "강서구", "기술개발"], ["3", "CC연구원", "전문연", "420", "충남", "천안시", "기반구축"]]
        cw = [(pw - 48) * w for w in (0.08, 0.17, 0.12, 0.13, 0.12, 0.2, 0.18)]
        ty = y + 24
        for r, row in enumerate(rows):
            x = px + 24
            for c, v in enumerate(row):
                d.rectangle([x, ty, x + cw[c], ty + 20], fill=(238, 240, 243) if r == 0 else "white", outline=(150, 150, 150))
                d.text((x + 3, ty + 3), v, font=font(9, r == 0), fill=INK)
                x += cw[c]
            ty += 20


def s5(n):
    im, d = base(4, "out/ 에 완성된 보고서", "결과")
    report_page(im, d, n)
    notes = ["템플릿 글꼴·크기 그대로", "날짜는 생성일로 자동", "자유 형식도 양식에 맞춤", "표는 엑셀 열·행 수만큼"]
    for i, t in enumerate(notes[: max(0, n - 4)]):
        cy = 175 + i * 52
        d.line([(50, cy), (58, cy + 8), (72, cy - 8)], fill=GREEN, width=4)
        d.text((80, 163 + i * 52), t, font=font(17), fill=INK)
    return im


staged(s5, 8, 450, 3800)


# ── 6. 입력 → 출력 대응 ───────────────────────────────
PAIRS = [
    ("1. 개요", "□ 추진 개요"),
    ("ㅇ 배경 : 정부 방침인 …", "ㅇ (배경) 정부 방침인 …"),
    ("* 국정과제 000 …  (- 단계 생략)", "   * 국정과제 000 …  (생략 그대로)"),
    ("(날짜 없음)", "26.10.8.(목), 대외협력실  (생성일)"),
    ("(참고자료 없음)", "참고1 문단 자동 삭제"),
    ("엑셀 6열 · 데이터 행", "표 열·행 자동 맞춤 + 순번"),
]


def s6(n):
    im, d = base(5, "사람이 쓴 그대로 → 양식에 맞게", "판단")
    d.text((60, 112), "입력 (내용파일)", font=font(15, True), fill=SUB)
    d.text((540, 112), "출력 (보고서)", font=font(15, True), fill=SUB)
    for i, (a, b) in enumerate(PAIRS[: min(n, len(PAIRS))]):
        y = 142 + i * 54
        box(d, [48, y, 440, y + 46], "white", LINE, 8, 1)
        d.text((64, y + 11), a, font=font(16), fill=INK)
        arrow(d, 452, y + 23, 516, BLUE)
        box(d, [528, y, W - 48, y + 46], LBLUE, None, 8)
        d.text((544, y + 11), b, font=font(16, True), fill=BLUE)
    if n > len(PAIRS):
        d.text((48, 468), "없는 내용은 지어내지 않고, 모호하면 사용자에게 먼저 묻습니다.", font=font(15), fill=SUB)
    return im


staged(s6, len(PAIRS) + 1, 650, 3800)


# ── 7. 사용법 · 마무리 ────────────────────────────────
def s7(n):
    im, d = base(6, "이렇게 씁니다", "사용법")
    steps = [
        ("1", "raw/ 에 템플릿 · 내용파일 · 엑셀(선택)을 넣는다"),
        ("2", "Claude Code에서  \"raw에 넣어둔 걸로 보고서 만들어줘\""),
        ("3", "out/ 의 결과를 한글에서 열어 확인 — 수정 0분이면 완료"),
    ]
    for i, (k, t) in enumerate(steps[: min(n, 3)]):
        y = 130 + i * 76
        d.ellipse([48, y, 92, y + 44], fill=BLUE)
        d.text((63, y + 6), k, font=font(22, True), fill="white")
        d.text((112, y + 8), t, font=font(20), fill=INK)
    if n >= 4:
        box(d, [48, 372, W - 48, 470], LBLUE, None)
        d.text((72, 386), "서비스(스크립트)로 먼저 만들고 → 에이전트(스킬)로 전환 →", font=font(18, True), fill=BLUE)
        d.text((72, 420), "실제로 돌려보며 발견한 문제를 git 커밋 단위로 검토하며 고쳤습니다.", font=font(18), fill=INK)
    return im


staged(s7, 4, 900, 4200)

# 저장 (프레임별 팔레트)
pal = [f.convert("P", palette=Image.ADAPTIVE, colors=128) for f in frames]
pal[0].save(OUT, save_all=True, append_images=pal[1:], duration=durs, loop=0, optimize=True, disposal=1)
print(len(frames), "frames,", sum(durs) / 1000, "s")
