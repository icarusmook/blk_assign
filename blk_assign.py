"""
blk_assign.py - 블록 작업장 배정 시스템

흐름:
  1. blk_master / area_capa 로드
  2. 각 blk의 load_mh를 작업기간(m_stdt~m_fndt) 내 영업일에 균등 분배 → 월별/주별 부하 산출
  3. 블록 배정 순서(우선순위)에 따라 순차 배정
     - 후보 작업장: area_prior_1~5
     - 배정 기준: 배정 후 해당 기간 조업도(부하/능력)가 가장 낮은(여유 있는) 작업장 선택
  4. assigned_area 컬럼으로 결과 저장
"""

import pandas as pd
import numpy as np
from datetime import date, timedelta
from collections import defaultdict
import openpyxl
from typing import Optional

# ── 경로 상수 ─────────────────────────────────────────────────────────────────
BASE_DIR        = '/mnt/d/mook/AI/pjt/company/blk_assign'
BLK_MASTER_PATH = f'{BASE_DIR}/blk_master.xlsx'
AREA_CAPA_PATH  = f'{BASE_DIR}/area_capa.xlsx'

# ── 공휴일 (2026 + 2027 일부) ──────────────────────────────────────────────────
KR_HOLIDAYS = {
    # 2026
    date(2026,  1,  1), date(2026,  2, 16), date(2026,  2, 17), date(2026,  2, 18),
    date(2026,  3,  1), date(2026,  5,  5), date(2026,  5, 26), date(2026,  6,  6),
    date(2026,  8, 15), date(2026,  9, 25), date(2026,  9, 26), date(2026,  9, 27),
    date(2026,  9, 28), date(2026, 10,  3), date(2026, 10,  9), date(2026, 12, 25),
    # 2027
    date(2027,  1,  1), date(2027,  1, 27), date(2027,  1, 28), date(2027,  1, 29),
    date(2027,  3,  1),
}

AREA_PRIOR_COLS = ['area_prior_1', 'area_prior_2', 'area_prior_3',
                   'area_prior_4', 'area_prior_5']

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 1. 영업일 유틸리티
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def is_workday(d: date) -> bool:
    return d.weekday() < 5 and d not in KR_HOLIDAYS


def get_workdays_between(start: date, end: date) -> list:
    """start ~ end 사이 영업일 목록"""
    days, cur = [], start
    while cur <= end:
        if is_workday(cur):
            days.append(cur)
        cur += timedelta(days=1)
    return days


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 2. 데이터 로드
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def load_blk_master(path: str) -> pd.DataFrame:
    df = pd.read_excel(path)
    df['m_stdt'] = pd.to_datetime(df['m_stdt'])
    df['m_fndt'] = pd.to_datetime(df['m_fndt'])
    return df


def load_area_capa(path: str) -> dict:
    """
    area_capa.xlsx 파싱
    반환: {area_name: {month_int: capacity_mh}}
    예:  {'AREA1': {1: 4035, 2: 3217, ..., 12: 4430}, ...}
    """
    wb   = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws   = wb.active
    rows = list(ws.iter_rows(values_only=True))
    # 행1: 연도 헤더, 행2: 월 헤더, 행3~: 데이터
    capa = {}
    for row in rows[2:]:
        area = row[0]
        if not area:
            continue
        capa[area] = {m: (row[m] or 0) for m in range(1, 13)}
    return capa


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 3. 부하 분산 계산 (핵심 공통 로직 — 모든 배정 단계에서 동일 적용)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def calc_load_distribution(row: pd.Series) -> dict:
    """
    load_mh를 m_stdt ~ m_fndt 기간의 영업일에 균등 분배.

    반환:
        {
          'monthly': {(year, month): load_mh, ...},
          'weekly' : {(iso_year, iso_week): load_mh, ...},
          'daily_mh': float  # 일 단위 부하
        }
    load_mh 또는 날짜가 없으면 빈 딕트 반환.
    """
    load_mh = row['load_mh']
    m_stdt  = row['m_stdt']
    m_fndt  = row['m_fndt']

    if pd.isna(load_mh) or pd.isna(m_stdt) or pd.isna(m_fndt):
        return {'monthly': {}, 'weekly': {}, 'daily_mh': 0.0}

    workdays = get_workdays_between(m_stdt.date(), m_fndt.date())
    if not workdays:
        return {'monthly': {}, 'weekly': {}, 'daily_mh': 0.0}

    daily_mh = load_mh / len(workdays)
    monthly: dict = defaultdict(float)
    weekly:  dict = defaultdict(float)

    for d in workdays:
        monthly[(d.year, d.month)] += daily_mh
        iso = d.isocalendar()
        weekly[(iso.year, iso.week)] += daily_mh

    return {
        'monthly' : dict(monthly),
        'weekly'  : dict(weekly),
        'daily_mh': daily_mh,
    }


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 4. 배정 알고리즘
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def get_candidate_areas(row_dict: dict, capa: dict) -> list:
    """
    area_prior_1~5 에서 유효 후보 작업장 추출
    - '동일' → m_area 값으로 대체
    - capa에 존재하지 않는 작업장 제외
    - 중복 제거 (순서 유지)
    """
    m_area = row_dict.get('m_area', '') or ''
    seen, result = set(), []

    for col in AREA_PRIOR_COLS:
        val = row_dict.get(col, '')
        if val is None or (isinstance(val, float) and np.isnan(val)) or val == '':
            continue
        if val == '동일':
            val = m_area
        if val and val in capa and val not in seen:
            seen.add(val)
            result.append(val)

    return result


def _area_util_score(
    area:         str,
    monthly_load: dict,            # 이 블록의 월별 부하
    current_load: dict,            # 현재까지 누적된 전체 부하 {area: {(yr,mo): load}}
    capa:         dict,
) -> float:
    """
    해당 작업장에 블록을 배정했을 때의 조업도 점수.
    블록이 걸치는 달들의 (기존부하 + 블록부하) / 능력 의 가중평균.
    낮을수록 여유 있음.
    """
    area_load = current_load[area]
    score_sum, weight_sum = 0.0, 0.0

    for (yr, mo), blk_load in monthly_load.items():
        cap = capa[area].get(mo, 0)
        if cap <= 0:
            continue
        new_load    = area_load.get((yr, mo), 0.0) + blk_load
        score_sum  += (new_load / cap) * blk_load   # 부하 가중
        weight_sum += blk_load

    return (score_sum / weight_sum) if weight_sum > 0 else 0.0


def best_area_for_block(
    candidates:   list,
    monthly_load: dict,
    current_load: dict,
    capa:         dict,
) -> Optional[str]:
    """후보 중 배정 후 조업도가 가장 낮은(여유 있는) 작업장 반환"""
    best_area, best_score = None, float('inf')

    for area in candidates:
        score = _area_util_score(area, monthly_load, current_load, capa)
        if score < best_score:
            best_score = score
            best_area  = area

    return best_area


# ── 배정 순서 기준 (추후 변경) ────────────────────────────────────────────────
# 현재: m_stdt 오름차순 → load_mh 내림차순 (착수 빠른 것, 부하 큰 것 우선)
# 변경 시 이 함수만 수정하면 됩니다.
def get_sort_order(df: pd.DataFrame) -> list:
    df_tmp = df.copy()
    df_tmp['_load_neg'] = -df_tmp['load_mh'].fillna(0)
    sorted_df = df_tmp.sort_values(['m_stdt', '_load_neg'], na_position='last')
    return sorted_df.index.tolist()


def assign_blocks(df: pd.DataFrame, capa: dict) -> tuple:
    """
    메인 배정 루프.
    반환: (결과 DataFrame, current_load dict, load_dists list)
    """
    # 3-1) 전체 부하 분산 사전 계산
    print("▶ 부하 분산 계산 중 ...")
    load_dists = [calc_load_distribution(row) for _, row in df.iterrows()]

    # 3-2) 배정 순서 결정
    sorted_idx = get_sort_order(df)

    # 3-3) 누적 부하 추적 테이블
    current_load: dict = {area: defaultdict(float) for area in capa}

    assigned_map: dict = {}

    print(f"▶ 배정 시작 (총 {len(sorted_idx):,}건) ...")
    for rank, orig_idx in enumerate(sorted_idx):
        row  = df.loc[orig_idx]
        dist = load_dists[orig_idx]
        monthly_load = dist['monthly']
        candidates   = get_candidate_areas(row.to_dict(), capa)

        if not candidates:
            assigned_map[orig_idx] = None
            continue

        if not monthly_load:
            # load_mh 없는 블록 → 1순위 작업장으로 직접 배정
            assigned_map[orig_idx] = candidates[0]
            continue

        area = best_area_for_block(candidates, monthly_load, current_load, capa)
        if area is None:
            area = candidates[0]

        assigned_map[orig_idx] = area

        # 누적 부하 업데이트
        for (yr, mo), load in monthly_load.items():
            current_load[area][(yr, mo)] += load

        if (rank + 1) % 1000 == 0:
            print(f"  {rank+1:,} / {len(sorted_idx):,} 완료")

    df = df.copy()
    df['assigned_area'] = pd.Series(assigned_map)
    return df, current_load, load_dists


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 5. 결과 요약 출력
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def print_utilization_table(current_load: dict, capa: dict):
    """작업장별 월별 조업도(%) 테이블 출력"""
    months = list(range(1, 13))

    print("\n" + "=" * 90)
    print("  작업장별 월별 조업도 (배정부하 / 능력 × 100%)")
    print("=" * 90)
    header = f"{'작업장':^8}" + "".join(f"  {m:>2}월" for m in months) + "  연평균"
    print(header)
    print("-" * len(header))

    all_utils = []
    for area in sorted(capa.keys()):
        row_utils = []
        for m in months:
            load = current_load[area].get((2026, m), 0.0)
            cap  = capa[area].get(m, 0)
            util = (load / cap * 100) if cap > 0 else 0.0
            row_utils.append(util)
        all_utils.append(row_utils)
        avg = sum(row_utils) / len(row_utils)
        print(f"{area:^8}" + "".join(f" {u:>5.1f}%" for u in row_utils) + f" {avg:>6.1f}%")

    print("-" * len(header))
    col_avgs = [sum(r[i] for r in all_utils) / len(all_utils) for i in range(12)]
    overall  = sum(col_avgs) / len(col_avgs)
    print(f"{'전체평균':^8}" + "".join(f" {u:>5.1f}%" for u in col_avgs) + f" {overall:>6.1f}%")


def print_monthly_load_summary(df: pd.DataFrame, load_dists: list, current_load: dict, capa: dict):
    """월별 배정 부하 합계 요약"""
    months = list(range(1, 13))

    print("\n" + "=" * 55)
    print("  월별 배정 부하 합계 (m_stdt 기준)")
    print("=" * 55)
    print(f"{'월':^8} {'블록수':>8} {'load_mh 합계':>14} {'능력 합계':>12} {'조업도':>8}")
    print("-" * 55)

    for m in months:
        blk_cnt  = ((df['m_stdt'].dt.month == m) & df['assigned_area'].notna()).sum()
        total_load = sum(current_load[a].get((2026, m), 0) for a in capa)
        total_cap  = sum(capa[a].get(m, 0) for a in capa)
        util = (total_load / total_cap * 100) if total_cap > 0 else 0
        print(f"2026-{m:02d}  {blk_cnt:>8,} {total_load:>14,.1f} {total_cap:>12,} {util:>7.1f}%")

    print("-" * 55)
    grand_load = sum(v for area in capa for v in current_load[area].values())
    grand_cap  = sum(capa[a].get(m, 0) for a in capa for m in months)
    grand_util = (grand_load / grand_cap * 100) if grand_cap > 0 else 0
    print(f"{'합계':^8} {df['assigned_area'].notna().sum():>8,} {grand_load:>14,.1f} {grand_cap:>12,} {grand_util:>7.1f}%")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 6. 저장
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def save_result(df: pd.DataFrame, path: str):
    df_out = df.copy()
    df_out['m_stdt'] = df_out['m_stdt'].dt.date
    df_out['m_fndt'] = df_out['m_fndt'].dt.date
    df_out.to_excel(path, index=False)
    print(f"\n✅ 저장 완료: {path}")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 메인
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

if __name__ == '__main__':
    print("=" * 55)
    print("  blk 작업장 배정 시스템")
    print("=" * 55)

    print("\n▶ 데이터 로드 중 ...")
    df   = load_blk_master(BLK_MASTER_PATH)
    capa = load_area_capa(AREA_CAPA_PATH)

    # assigned_area 컬럼이 이미 있으면 초기화
    if 'assigned_area' in df.columns:
        df['assigned_area'] = None

    print(f"  blk_master : {len(df):,}행 × {len(df.columns)}열")
    print(f"  area_capa  : {len(capa)}개 작업장 ({', '.join(sorted(capa.keys()))})")

    df, current_load, load_dists = assign_blocks(df, capa)

    print(f"\n▶ 배정 결과:")
    print(f"  배정 완료 : {df['assigned_area'].notna().sum():,}건")
    print(f"  미배정    : {df['assigned_area'].isna().sum():,}건")
    print(f"  작업장 분포:")
    dist = df['assigned_area'].value_counts().sort_index()
    for area, cnt in dist.items():
        print(f"    {area}: {cnt:,}건")

    print_monthly_load_summary(df, load_dists, current_load, capa)
    print_utilization_table(current_load, capa)

    save_result(df, BLK_MASTER_PATH)
