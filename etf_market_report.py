# -*- coding: utf-8 -*-
"""
국내 ETF 시장 데이터 자동 수집 · 검증 · 일일 리포트
(공공데이터포털 '금융위원회_증권상품시세정보' Open API)

무엇을 하나요
  1) 상장 ETF 전 종목의 일별 시세(종가, NAV, 거래대금, 순자산총액, 기초지수)를 수집합니다.
  2) 종목명 앞 브랜드(TIGER, KODEX, ACE ...)로 운용사를 구분합니다.
     브랜드 이름이 바뀐 경우(KBSTAR→RISE, ARIRANG→PLUS, KOSEF→KIWOOM)도 같은 운용사로 묶습니다.
  3) 데이터 이상을 자동 점검합니다(중복, 순자산 누락, NAV 괴리율, 수집 누락일, 미분류 브랜드).
  4) 운용사별 순자산 점유율, 월말 점유율 추이, 신규 상장, 테마별 자금 흐름을 Excel로 만듭니다.

자동화
  - 처음 실행하면 START_DATE부터 전부 받고, 원본을 etf_raw.csv에 저장합니다.
  - 다음부터는 마지막 수집일 이후 날짜만 추가로 받아 리포트를 다시 만듭니다.
  - Windows '작업 스케줄러'에 등록하면 매일 아침 리포트가 자동 갱신됩니다.

실행 방법
  1) pip install requests pandas openpyxl
  2) python etf_market_report.py
     처음 실행 시 공공데이터포털 인증키(Decoding 키)를 입력하라는 안내가 나옵니다.
     환경변수 DATA_GO_KR_KEY에 넣어 두면 묻지 않습니다(자동 실행 시 필요).
결과: 같은 폴더에 ETF_시장리포트_YYYYMMDD.xlsx
"""
import os
import sys
import time
import datetime as dt
from pathlib import Path

import requests
import pandas as pd

# ------------------------------------------------------------------ 설정
BASE_DIR = Path(__file__).resolve().parent
RAW_CSV = BASE_DIR / "etf_raw.csv"
START_DATE = "20250101"   # 처음 수집 시작일
ENDPOINT = "https://apis.data.go.kr/1160100/GetSecuritiesProductInfoService_V2"
# V2 기능 이름이 문서마다 달라서, 처음 호출할 때 응답이 오는 이름을 자동으로 찾습니다.
OPERATIONS = ["getETFPriceInfo_V2", "getETFPriceInfoV2", "getETFPriceInfo"]
URL = None
ROWS = 1000               # 한 번에 받을 행 수
SLEEP = 0.2               # 호출 간 대기(초)
PREMIUM_LIMIT = 3.0       # NAV 대비 괴리율 경고 기준(%)
OVERSEAS_KW = ["차이나", "중국", "항셍", "HSCEI", "CSI", "과창판", "심천", "ChiNext", "샤오미",
               "홍콩", "일본", "니케이", "인도", "베트남", "유럽", "러시아", "글로벌", "미국", "S&P",
               "나스닥", "NASDAQ", "다우", "MSCI World"]

# 브랜드 → 운용사 (옛 브랜드명 포함). 긴 이름을 먼저 검사합니다.
BRANDS = {
    "TIGER": "미래에셋자산운용",
    "KODEX": "삼성자산운용", "KoAct": "삼성액티브자산운용",
    "ACE": "한국투자신탁운용", "KINDEX": "한국투자신탁운용",
    "RISE": "KB자산운용", "KBSTAR": "KB자산운용",
    "SOL": "신한자산운용",
    "PLUS": "한화자산운용", "ARIRANG": "한화자산운용",
    "KIWOOM": "키움투자자산운용", "KOSEF": "키움투자자산운용",
    "HANARO": "NH-Amundi자산운용",
    "1Q": "하나자산운용",
    "TIMEFOLIO": "타임폴리오자산운용",
    "WON": "우리자산운용", "WOORI": "우리자산운용",
    "BNK": "BNK자산운용",
    "TRUSTON": "트러스톤자산운용",
    "UNICORN": "현대자산운용",
    "마이티": "DB자산운용",
    "에셋플러스": "에셋플러스자산운용",
    "파워": "교보악사자산운용",
    "히어로즈": "키움투자자산운용",
    "흥국": "흥국자산운용",
    "VITA": "한국투자밸류자산운용",
    "FOCUS": "브이아이자산운용",
    "DAISHIN343": "대신자산운용",
    "ITF": "IBK자산운용",
    # 2026-10 실제 데이터 검증에서 '미분류'로 나온 브랜드 추가 (새 브랜드명·리브랜딩 반영)
    "TIME": "타임폴리오자산운용",      # TIMEFOLIO → TIME
    "IBK": "IBK자산운용",
    "HK": "흥국자산운용",
    "MIDAS": "마이다스에셋자산운용", "마이다스": "마이다스에셋자산운용",
    "DAISHIN": "대신자산운용",
    "DS": "DS자산운용",
    "아이엠에셋": "iM에셋자산운용",
    "TREX": "유리자산운용",
    "KCGI": "KCGI자산운용",
    "더제이": "더제이자산운용",
}

# 기초지수명·종목명 키워드 → 테마 (위에서부터 먼저 맞는 것)
THEMES = [
    # 주식+채권을 섞은 상품은 '채권'만 보고 분류하면 안 되므로 먼저 분리합니다.
    # (주식 비중 50% 이하 채권혼합형은 퇴직연금 안전자산 한도로 투자할 수 있어 연금 자금이 몰리는 영역)
    ("채권혼합(연금)", ["채권혼합", "미국채혼합", "국채혼합"]),
    ("채권·금리", ["채권", "국채", "회사채", "CD금리", "KOFR", "SOFR", "머니마켓", "단기", "금리"]),
    ("금·원자재", ["금현물", "골드", "금선물", "은선물", "원유", "구리", "원자재"]),
    ("반도체", ["반도체", "Semiconductor", "SEMICONDUCTOR"]),
    ("AI·빅테크", ["AI", "인공지능", "빅테크", "테크TOP", "나스닥100", "NASDAQ"]),
    ("2차전지", ["2차전지", "이차전지", "배터리", "리튬"]),
    ("배당·인컴", ["배당", "커버드콜", "인컴", "리츠", "REIT"]),
    ("미국 대표지수", ["S&P", "미국", "다우"]),
    ("국내 대표지수", ["코스피", "KOSPI", "코스닥", "KOSDAQ", "KRX300", "200"]),
]


# ------------------------------------------------------------------ 수집
def get_key():
    key = os.environ.get("DATA_GO_KR_KEY", "").strip()
    if not key:
        key = input("공공데이터포털 인증키(Decoding)를 입력하세요: ").strip()
    if not key:
        sys.exit("인증키가 없습니다. data.go.kr > 마이페이지 > 개발계정에서 확인하세요.")
    return key


def fetch_page(key, begin, end, page):
    """endBasDt는 '미만' 조건이라 end 당일은 포함되지 않습니다."""
    params = dict(serviceKey=key, resultType="json", numOfRows=ROWS, pageNo=page,
                  beginBasDt=begin, endBasDt=end)
    global URL
    if URL is None:
        tried = []
        for op in OPERATIONS:
            r = requests.get(f"{ENDPOINT}/{op}", params=params, timeout=60)
            tried.append(f"{op} → {r.status_code}")
            if r.status_code == 200 and '"response"' in r.text:
                URL = f"{ENDPOINT}/{op}"
                print(f"API 주소 확인: {URL}")
                break
        else:
            sys.exit("\n[오류] ETF 시세 기능 주소를 찾지 못했습니다.\n시도 결과: " + " / ".join(tried) +
                     "\n→ 401이면 인증키 문제, 404면 기능 이름 문제입니다. 이 메시지를 그대로 보내 주세요."
                     f"\n마지막 응답 앞부분: {r.text[:200]}")
    else:
        r = requests.get(URL, params=params, timeout=60)
    if r.status_code != 200:
        hint = {
            401: "인증키가 아직 등록되지 않았거나(신청 후 1~2시간 소요) Encoding 키를 넣은 경우입니다.",
            403: "이 API에 대한 활용신청이 승인되지 않았습니다. data.go.kr 마이페이지에서 확인하세요.",
            429: "일일 호출 한도를 넘었습니다. 내일 다시 실행하세요.",
        }.get(r.status_code, "서버 일시 오류일 수 있습니다. 잠시 후 다시 실행하세요.")
        sys.exit(f"\n[오류] 서버 응답 코드 {r.status_code}\n→ {hint}\n응답 앞부분: {r.text[:200]}")
    time.sleep(SLEEP)
    try:
        body = r.json()["response"]["body"]
    except (ValueError, KeyError):
        sys.exit("응답을 읽지 못했습니다. 인증키(Decoding 키)가 맞는지, 활용신청 승인 후 "
                 "1~2시간이 지났는지 확인하세요.\n응답 앞부분: " + r.text[:300])
    items = body.get("items") or {}
    rows = items.get("item", []) if isinstance(items, dict) else []
    if isinstance(rows, dict):
        rows = [rows]
    return rows, int(body.get("totalCount", 0))


def fetch_range(key, begin, end):
    out, page = [], 1
    while True:
        rows, total = fetch_page(key, begin, end, page)
        out.extend(rows)
        print(f"  {begin}~{end} 페이지 {page}: 누적 {len(out):,} / {total:,}행")
        if not rows or len(out) >= total:
            break
        page += 1
    return out


def collect(key):
    """etf_raw.csv에 없는 날짜만 받아 합칩니다."""
    tomorrow = (dt.date.today() + dt.timedelta(days=1)).strftime("%Y%m%d")
    old = pd.DataFrame()
    begin = START_DATE
    if RAW_CSV.exists():
        old = pd.read_csv(RAW_CSV, dtype=str, encoding="utf-8-sig")
        last = old["basDt"].max()
        begin = (dt.datetime.strptime(last, "%Y%m%d") + dt.timedelta(days=1)).strftime("%Y%m%d")
        print(f"기존 데이터 마지막 날짜 {last} → {begin}부터 추가 수집")
    if begin >= tomorrow:
        print("새로 받을 날짜가 없습니다.")
        return old

    # 한 번에 너무 많은 페이지가 쌓이지 않게 달 단위로 나눠 받습니다.
    new_rows = []
    cur = dt.datetime.strptime(begin, "%Y%m%d").date()
    stop = dt.datetime.strptime(tomorrow, "%Y%m%d").date()
    while cur < stop:
        nxt = min((cur.replace(day=1) + dt.timedelta(days=32)).replace(day=1), stop)
        new_rows += fetch_range(key, cur.strftime("%Y%m%d"), nxt.strftime("%Y%m%d"))
        cur = nxt

    new = pd.DataFrame(new_rows).astype(str)
    df = pd.concat([old, new], ignore_index=True) if len(old) else new
    df.to_csv(RAW_CSV, index=False, encoding="utf-8-sig")
    print(f"원본 저장: {RAW_CSV.name} ({len(df):,}행, 신규 {len(new):,}행)")
    return df


# ------------------------------------------------------------------ 정리
NUM_COLS = ["clpr", "nav", "trqu", "trPrc", "mrktTotAmt", "nPptTotAmt", "stLstgCnt", "fltRt"]


def brand_of(name):
    for b in sorted(BRANDS, key=len, reverse=True):
        if str(name).upper().startswith(b.upper()):
            return b
    return str(name).split()[0] if str(name).strip() else ""


def theme_of(row):
    text = f"{row.get('bssIdxIdxNm', '')} {row.get('itmsNm', '')}"
    for theme, kws in THEMES:
        if any(k in text for k in kws):
            return theme
    return "기타"


def prepare(df):
    df = df.copy()
    need = ["basDt", "srtnCd", "itmsNm", "clpr", "nav", "nPptTotAmt", "trPrc"]
    missing = [c for c in need if c not in df.columns]
    if missing:
        sys.exit(f"\n[오류] V2 응답에 없는 항목: {missing}\n실제 항목: {list(df.columns)}"
                 "\n→ 이 메시지를 그대로 보내 주시면 항목 이름을 맞춰 드리겠습니다.")
    if "bssIdxIdxNm" not in df.columns:
        df["bssIdxIdxNm"] = ""
    for c in NUM_COLS:
        if c in df:
            df[c] = pd.to_numeric(df[c].astype(str).str.replace(",", ""), errors="coerce")
    df["date"] = pd.to_datetime(df["basDt"], format="%Y%m%d")
    df["브랜드"] = df["itmsNm"].map(brand_of)
    df["운용사"] = df["브랜드"].map(BRANDS).fillna("미분류")
    df["테마"] = df.apply(theme_of, axis=1)
    df["괴리율(%)"] = (df["clpr"] / df["nav"] - 1) * 100
    df["순자산(억원)"] = df["nPptTotAmt"] / 1e8
    return df


# ------------------------------------------------------------------ 검증
def validate(df):
    checks = []

    def add(name, n, detail, rule):
        checks.append({"점검 항목": name, "기준": rule, "건수": n,
                       "결과": "정상" if n == 0 else "확인 필요", "내용(최대 10건)": detail})

    dup = df[df.duplicated(["basDt", "srtnCd"], keep=False)]
    add("날짜·종목코드 중복", len(dup), ", ".join(dup["itmsNm"].head(10)),
        "같은 날 같은 종목이 두 번 이상")

    miss = df[df["nPptTotAmt"].isna() | (df["nPptTotAmt"] <= 0)]
    add("순자산총액 누락·0", len(miss), ", ".join((miss["basDt"] + " " + miss["itmsNm"]).head(10)),
        "nPptTotAmt가 비었거나 0 이하")

    # 괴리율 경고를 원인별로 나눕니다(1차 실행에서 6,420건이 나와 규칙을 보완).
    prem = df[df["괴리율(%)"].abs() > PREMIUM_LIMIT].copy()
    overseas = prem["itmsNm"].str.contains("|".join(OVERSEAS_KW)) | \
        prem["bssIdxIdxNm"].astype(str).str.contains("|".join(OVERSEAS_KW))
    halted = prem["trPrc"].fillna(0) <= 0
    prem["원인"] = "국내 기초·정상 거래(실제 점검 대상)"
    prem.loc[overseas, "원인"] = "해외 기초(현지 휴장·시차로 NAV 반영 지연 가능)"
    prem.loc[halted, "원인"] = "거래대금 0(거래정지 의심)"

    def fmt(x):
        return ", ".join((x["basDt"] + " " + x["itmsNm"] + " " +
                          x["괴리율(%)"].round(1).astype(str) + "%").head(10))

    add(f"NAV 괴리율 ±{PREMIUM_LIMIT}% 초과(전체)", len(prem),
        " / ".join(f"{k} {v:,}건" for k, v in prem["원인"].value_counts().items()), "종가 ÷ NAV − 1")
    for cause in ["거래대금 0(거래정지 의심)", "해외 기초(현지 휴장·시차로 NAV 반영 지연 가능)",
                  "국내 기초·정상 거래(실제 점검 대상)"]:
        sub = prem[prem["원인"] == cause]
        add(f"  └ {cause}", len(sub), fmt(sub.sort_values("괴리율(%)", key=abs, ascending=False)),
            "위 경고 중 해당 원인")
    top_days = prem[overseas & ~halted].groupby("basDt").size().sort_values(ascending=False).head(10)
    add("  └ 해외 기초 경고가 몰린 날(휴장일 확인용)", len(top_days),
        ", ".join(f"{d}({n}건)" for d, n in top_days.items()), "현지 연휴와 겹치는지 확인")

    cnt = df.groupby("basDt")["srtnCd"].nunique()
    drop = cnt[cnt < cnt.rolling(5, min_periods=1).median() * 0.9]
    add("종목 수 급감일(수집 누락 의심)", len(drop),
        ", ".join(f"{d}({n}종목)" for d, n in drop.head(10).items()),
        "그날 종목 수 < 최근 5일 중앙값의 90%")

    unk = df[df["운용사"] == "미분류"].drop_duplicates("srtnCd")
    add("운용사 미분류 브랜드", len(unk), ", ".join(unk["itmsNm"].head(10)),
        "BRANDS 사전에 없는 브랜드 → 사전에 추가")

    latest = df[df["basDt"] == df["basDt"].max()]
    share_sum = (latest.groupby("운용사")["nPptTotAmt"].sum() / latest["nPptTotAmt"].sum()).sum()
    add("최신일 점유율 합계 = 100%", int(abs(share_sum - 1) > 1e-6), f"합계 {share_sum*100:.4f}%",
        "운용사별 점유율을 모두 더한 값")
    return pd.DataFrame(checks)


# ------------------------------------------------------------------ 분석
def analyze(df):
    latest_d = df["date"].max()
    latest = df[df["date"] == latest_d]

    # 월말 기준일
    month_end = df.groupby(df["date"].dt.to_period("M"))["date"].max()
    prev_me = month_end[month_end < latest_d.to_period("M").start_time]
    prev_d = prev_me.iloc[-1] if len(prev_me) else df["date"].min()
    prev = df[df["date"] == prev_d]

    def share(x):
        g = x.groupby("운용사").agg(순자산=("순자산(억원)", "sum"), 종목수=("srtnCd", "nunique"),
                                  거래대금=("trPrc", "sum"))
        g["점유율(%)"] = g["순자산"] / g["순자산"].sum() * 100
        return g

    now, bef = share(latest), share(prev)
    summary = now.join(bef[["순자산", "점유율(%)"]], rsuffix="_전월말").fillna(0)
    summary["순자산 증감(억원)"] = summary["순자산"] - summary["순자산_전월말"]
    summary["점유율 변화(%p)"] = summary["점유율(%)"] - summary["점유율(%)_전월말"]
    summary["거래대금(억원)"] = summary["거래대금"] / 1e8
    summary = summary.sort_values("순자산", ascending=False).reset_index()
    summary.insert(0, "순위", range(1, len(summary) + 1))
    summary = summary[["순위", "운용사", "종목수", "순자산", "점유율(%)", "순자산 증감(억원)",
                       "점유율 변화(%p)", "거래대금(억원)"]].rename(columns={"순자산": "순자산(억원)"})

    # 월말 점유율 추이 (상위 6개사 + 기타)
    me = df[df["date"].isin(month_end.values)]
    top = summary["운용사"].head(6).tolist()
    me = me.assign(구분=me["운용사"].where(me["운용사"].isin(top), "기타"))
    trend = me.pivot_table(index=me["date"].dt.strftime("%Y-%m"), columns="구분",
                           values="순자산(억원)", aggfunc="sum")
    trend = trend.div(trend.sum(axis=1), axis=0) * 100
    trend = trend[[c for c in top + ["기타"] if c in trend.columns]].round(2)
    trend.index.name = "월"

    # 신규 상장: 수집 첫날 이후 처음 등장한 종목
    first = df.groupby("srtnCd")["date"].min()
    new_codes = first[first > df["date"].min()]
    new = (df[df["srtnCd"].isin(new_codes.index)].sort_values("date")
           .drop_duplicates("srtnCd")[["srtnCd", "itmsNm", "운용사", "테마", "bssIdxIdxNm", "date"]])
    new = new.merge(latest[["srtnCd", "순자산(억원)"]], on="srtnCd", how="left")
    new = new.rename(columns={"srtnCd": "종목코드", "itmsNm": "종목명", "bssIdxIdxNm": "기초지수",
                              "date": "첫 거래일", "순자산(억원)": "현재 순자산(억원)"})
    new["첫 거래일"] = new["첫 거래일"].dt.strftime("%Y-%m-%d")
    new = new.sort_values("첫 거래일", ascending=False)

    # 테마별 순자산: 전체 vs TIGER (전월말 → 최신일)
    def theme_tbl(x):
        return x.groupby("테마")["순자산(억원)"].sum()

    th = pd.DataFrame({
        "시장 전체(전월말)": theme_tbl(prev), "시장 전체(최신)": theme_tbl(latest),
        "TIGER(전월말)": theme_tbl(prev[prev["브랜드"] == "TIGER"]),
        "TIGER(최신)": theme_tbl(latest[latest["브랜드"] == "TIGER"]),
    }).fillna(0)
    th["시장 증감(억원)"] = th["시장 전체(최신)"] - th["시장 전체(전월말)"]
    th["TIGER 점유율(%)"] = th["TIGER(최신)"] / th["시장 전체(최신)"].where(th["시장 전체(최신)"] > 0) * 100
    th = th.sort_values("시장 전체(최신)", ascending=False).round(1).reset_index()

    # TIGER 순자산 상위 20
    tiger = (latest[latest["브랜드"] == "TIGER"]
             .sort_values("순자산(억원)", ascending=False).head(20)
             [["srtnCd", "itmsNm", "테마", "순자산(억원)", "clpr", "nav", "괴리율(%)", "trPrc"]]
             .rename(columns={"srtnCd": "종목코드", "itmsNm": "종목명", "clpr": "종가", "nav": "NAV",
                              "trPrc": "거래대금(원)"}))

    meta = {"최신일": latest_d.strftime("%Y-%m-%d"), "비교일(전월말)": prev_d.strftime("%Y-%m-%d"),
            "수집 기간": f"{df['date'].min():%Y-%m-%d} ~ {latest_d:%Y-%m-%d}",
            "거래일 수": df["basDt"].nunique(), "최신일 상장 ETF": latest["srtnCd"].nunique(),
            "최신일 시장 순자산(조원)": round(latest["순자산(억원)"].sum() / 1e4, 1)}
    return meta, summary, trend, new, th, tiger, latest


# ------------------------------------------------------------------ Excel
def write_excel(meta, summary, trend, new, th, tiger, latest, checks):
    from openpyxl.chart import LineChart, Reference
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    out = BASE_DIR / f"ETF_시장리포트_{meta['최신일'].replace('-', '')}.xlsx"
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        pd.DataFrame(list(meta.items()), columns=["항목", "값"]).to_excel(xw, sheet_name="요약", index=False)
        summary.to_excel(xw, sheet_name="요약", index=False, startrow=len(meta) + 3)
        trend.to_excel(xw, sheet_name="월말점유율추이")
        th.to_excel(xw, sheet_name="테마별자금흐름", index=False)
        new.to_excel(xw, sheet_name="신규상장", index=False)
        tiger.to_excel(xw, sheet_name="TIGER상위20", index=False)
        checks.to_excel(xw, sheet_name="데이터검증", index=False)
        raw = latest.drop(columns=["date"]).sort_values("nPptTotAmt", ascending=False)
        raw.to_excel(xw, sheet_name="최신일원본", index=False)

        wb = xw.book
        head_fill = PatternFill("solid", fgColor="1F3864")
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, float):
                        c.number_format = "#,##0.0"
            for i, col in enumerate(ws.columns, 1):
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[get_column_letter(i)].width = min(max(10, width * 1.3), 50)
        for ws in wb.worksheets:
            hdr = len(meta) + 4 if ws.title == "요약" else 1
            for c in ws[hdr]:
                c.font, c.fill = Font(bold=True, color="FFFFFF"), head_fill
                c.alignment = Alignment(horizontal="center")
        ws = wb["요약"]
        ws.insert_rows(1)
        ws["A1"] = "국내 ETF 시장 일일 리포트 (자동 생성)"
        ws["A1"].font = Font(bold=True, size=14)

        ws = wb["월말점유율추이"]
        ch = LineChart()
        ch.title, ch.y_axis.title, ch.height, ch.width = "운용사별 순자산 점유율(%)", "%", 9, 22
        ch.add_data(Reference(ws, min_col=2, max_col=ws.max_column, min_row=1, max_row=ws.max_row),
                    titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=ws.max_row))
        ws.add_chart(ch, f"{get_column_letter(ws.max_column + 2)}2")

        ws = wb["데이터검증"]
        warn = PatternFill("solid", fgColor="FCE4D6")
        for r in range(2, ws.max_row + 1):
            if ws.cell(r, 4).value == "확인 필요":
                for c in ws[r]:
                    c.fill = warn
    return out


def main():
    key = get_key()
    raw = collect(key)
    if raw.empty:
        sys.exit("수집된 데이터가 없습니다.")
    df = prepare(raw)
    checks = validate(df)
    meta, summary, trend, new, th, tiger, latest = analyze(df)
    out = write_excel(meta, summary, trend, new, th, tiger, latest, checks)
    print("\n[검증 결과]")
    print(checks[["점검 항목", "건수", "결과"]].to_string(index=False))
    print(f"\n완료: {out.name}")


if __name__ == "__main__":
    main()
