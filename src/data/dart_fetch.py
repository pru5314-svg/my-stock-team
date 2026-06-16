# -*- coding: utf-8 -*-
"""DART(전자공시) 데이터 수집 — 삼성전자 공시·재무.

사용:
    python src/data/dart_fetch.py
환경변수 DART_KEY 필요 (.env).
"""
import os
import sys
import json
import io
from pathlib import Path
from datetime import datetime, timedelta

import requests

# Windows 콘솔 UTF-8 출력
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

ROOT = Path(__file__).resolve().parents[2]
RAW = ROOT / "data" / "raw"
RAW.mkdir(parents=True, exist_ok=True)


def load_key():
    # .env 직접 파싱 (의존성 없이)
    env = ROOT / ".env"
    if env.exists():
        for line in env.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())
    key = os.environ.get("DART_KEY")
    if not key:
        sys.exit("DART_KEY 가 .env 또는 환경변수에 없습니다.")
    return key


KEY = load_key()
CORP = "00126380"        # 삼성전자 DART 고유번호
STOCK = "005930"
BASE = "https://opendart.fss.or.kr/api"


def get(path, **params):
    params["crtfc_key"] = KEY
    r = requests.get(f"{BASE}/{path}", params=params, timeout=20)
    r.raise_for_status()
    return r.json()


def recent_disclosures(days=30):
    end = datetime(2026, 6, 16)
    start = end - timedelta(days=days)
    data = get(
        "list.json",
        corp_code=CORP,
        bgn_de=start.strftime("%Y%m%d"),
        end_de=end.strftime("%Y%m%d"),
        page_count=100,
    )
    (RAW / "samsung_disclosures.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return data


def financials(year, reprt_code):
    # reprt_code: 11011=사업(연간) 11013=1Q 11012=반기 11014=3Q
    data = get(
        "fnlttSinglAcntAll.json",
        corp_code=CORP,
        bsns_year=str(year),
        reprt_code=reprt_code,
        fs_div="CFS",   # 연결재무제표
    )
    fn = RAW / f"samsung_fin_{year}_{reprt_code}.json"
    fn.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return data


def won(v):
    try:
        return f"{int(v):,}"
    except (ValueError, TypeError):
        return v


def pick(items, names):
    """account_nm 이 names 중 하나인 첫 항목 반환."""
    for it in items:
        if it.get("account_nm") in names:
            return it
    return None


def main():
    print("=" * 60)
    print("삼성전자(005930) DART 공시·재무 수집")
    print("=" * 60)

    # 1) 최근 공시
    dis = recent_disclosures(45)
    print(f"\n[최근 공시] status={dis.get('status')} {dis.get('message','')}")
    rows = dis.get("list", [])
    print(f"건수: {len(rows)}")
    for it in rows[:15]:
        print(f"  {it['rcept_dt']}  {it['report_nm']}  ({it['flr_nm']})")

    # 2) 재무제표 — 최근 분기/연간 시도
    print("\n[재무제표]")
    targets = [(2026, "11013"), (2025, "11011"), (2024, "11011")]
    for year, code in targets:
        try:
            fin = financials(year, code)
        except Exception as e:
            print(f"  {year}/{code} 실패: {e}")
            continue
        if fin.get("status") != "000":
            print(f"  {year}/{code}: {fin.get('status')} {fin.get('message')}")
            continue
        items = fin.get("list", [])
        label = {"11011": "연간", "11013": "1Q", "11012": "반기", "11014": "3Q"}[code]
        print(f"\n  === {year} {label} (연결, 항목 {len(items)}) ===")
        # 손익계산서/재무상태표 주요 계정
        keys = [
            (["매출액", "수익(매출액)", "영업수익"], "매출액"),
            (["영업이익", "영업이익(손실)"], "영업이익"),
            (["당기순이익", "당기순이익(손실)"], "당기순이익"),
            (["자산총계"], "자산총계"),
            (["부채총계"], "부채총계"),
            (["자본총계"], "자본총계"),
        ]
        for names, label2 in keys:
            it = pick(items, names)
            if it:
                cur = won(it.get("thstrm_amount"))
                print(f"    {label2:8} {cur} 원  (계정:{it.get('account_nm')})")
        # 첫 성공 연간 데이터까지만 상세 출력하면 충분
    print("\n원본 JSON: data/raw/ 에 저장됨")


if __name__ == "__main__":
    main()
