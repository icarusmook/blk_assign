"""
blk_assign_agent.py — 블록 작업장 배정 시스템

[글로벌 규칙]
  · load_mh를 m_stdt~m_fndt 영업일 균등 분배 → 월별/주별 부하 산출 (전 규칙 공통)
  · 전체 조업도(부하/능력) 균등화 목표 / 인접 주 여유도 반영
  · 결과: assigned_area 컬럼 추가 → blk_assign_result.xlsx

[순차 배정 규칙]
  R0.  이관물량(H) 처리가능여부 리포팅
  R1.  작업장별 월별 배정 목표물량 산출
  R2.  H/Y9 블록 → AREA15 (목표물량 무관)
  R3.  T-중/대   → m_area
  R4.  T-소/동일 → prt_area
  R5.  H-소 AREA3~5 처리가능여부 리포팅
  R6.  H-소 → AREA3~5 순환(round-robin)
  R7.  H-중-jig_F → AREA2 (Part1 특정blk 무조건 + Part2 목표까지)
  R8.  H-중-jig_D/L/W → AREA1→2→6 순환
  R9.  H-전문화 family → AREA7/8/9/10
  R10. H-대-jig_D/L 그룹 → AREA12
  R11. H-대/중 잔여 → AREA7~14 (PS페어 우선, 불가시 개별)
  R12. H-dir(blk[3]=9/H) → prt_area (override)
  R13. T-소 분산배정: area_seq 순, (pjt+blk+type) 그룹, assign_prior→m_stdt 정렬
"""

import json
import pandas as pd
import numpy as np
from datetime import date, timedelta
from collections import defaultdict
import openpyxl

# ── 경로 ─────────────────────────────────────────────────────────────────────
BASE_DIR        = '/mnt/d/mook/AI/pjt/company/blk_assign'
BLK_MASTER_PATH = f'{BASE_DIR}/blk_master.xlsx'
AREA_CAPA_PATH  = f'{BASE_DIR}/area_capa.xlsx'
RESULT_PATH     = f'{BASE_DIR}/blk_assign_result.xlsx'
STATS_PATH      = f'{BASE_DIR}/blk_assign_stats.json'

# ── 규칙 추적 ─────────────────────────────────────────────────────────────────
_current_rule: str = ''

def set_rule(rule: str):
    global _current_rule
    _current_rule = rule

# ── 공휴일 ───────────────────────────────────────────────────────────────────
KR_HOLIDAYS = {
    date(2026,  1,  1),
    date(2026,  2, 16), date(2026,  2, 17), date(2026,  2, 18),
    date(2026,  3,  1),
    date(2026,  5,  5), date(2026,  5, 26),
    date(2026,  6,  6),
    date(2026,  8, 15),
    date(2026,  9, 25), date(2026,  9, 26), date(2026,  9, 27), date(2026,  9, 28),
    date(2026, 10,  3), date(2026, 10,  9),
    date(2026, 12, 25),
    date(2027,  1,  1), date(2027,  1, 27), date(2027,  1, 28), date(2027,  1, 29),
    date(2027,  3,  1),
}

# ── 작업장 그룹 상수 ──────────────────────────────────────────────────────────
AREA_Y9       = 'AREA15'
AREAS_SMALL   = ['AREA3', 'AREA4', 'AREA5']
AREAS_MID_DL  = ['AREA1', 'AREA2', 'AREA6']
AREAS_REMAIN  = ['AREA7', 'AREA8', 'AREA9', 'AREA10', 'AREA12', 'AREA13', 'AREA14']
PS_AREA_PAIRS = [('AREA7', 'AREA8'), ('AREA9', 'AREA10'), ('AREA13', 'AREA14')]


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 유틸리티
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def is_workday(d: date) -> bool:
    return d.weekday() < 5 and d not in KR_HOLIDAYS


def get_workdays_between(start: date, end: date) -> list:
    days, cur = [], start
    while cur <= end:
        if is_workday(cur):
            days.append(cur)
        cur += timedelta(days=1)
    return days


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 데이터 로드
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def load_blk_master(path: str) -> pd.DataFrame:
    df = pd.read_excel(path)
    df['m_stdt'] = pd.to_datetime(df['m_stdt'])
    df['m_fndt'] = pd.to_datetime(df['m_fndt'])
    df['assigned_area'] = None
    df['rule_assigned']  = ''
    return df


def load_area_capa(path: str) -> dict:
    """반환: {area_name: {month_int: capacity_mh}}"""
    wb   = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws   = wb.active
    rows = list(ws.iter_rows(values_only=True))
    capa = {}
    for row in rows[2:]:
        area = row[0]
        if not area:
            continue
        capa[area] = {m: (row[m] or 0) for m in range(1, 13)}
    return capa


def load_area_seq(path: str) -> list:
    """area_capa.xlsx col[13](assign_prior) 오름차순으로 정렬된 area 목록 반환"""
    wb   = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws   = wb.active
    rows = list(ws.iter_rows(values_only=True))
    areas = []
    for row in rows[2:]:
        if row[0]:
            seq = row[13] if (len(row) > 13 and row[13] is not None) else 999
            areas.append((int(seq), row[0]))
    return [name for _, name in sorted(areas)]


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 글로벌: 부하 분산 계산 (모든 배정 규칙에 공통 적용)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def calc_load_distribution(row: pd.Series) -> dict:
    """
    load_mh를 m_stdt~m_fndt 영업일에 균등 분배.
    반환: {monthly: {(yr,mo): mh}, weekly: {(iso_yr,iso_wk): mh}, daily_mh: float}
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
    monthly, weekly = defaultdict(float), defaultdict(float)
    for d in workdays:
        monthly[(d.year, d.month)] += daily_mh
        iso = d.isocalendar()
        weekly[(iso.year, iso.week)] += daily_mh
    return {'monthly': dict(monthly), 'weekly': dict(weekly), 'daily_mh': daily_mh}


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 글로벌 헬퍼: 부하 추적 & 배정
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _init_area(area: str, cl: dict, cwl: dict):
    if area not in cl:
        cl[area]  = defaultdict(float)
        cwl[area] = defaultdict(float)


def add_load(area: str, dist: dict, cl: dict, cwl: dict):
    _init_area(area, cl, cwl)
    for k, v in dist['monthly'].items():
        cl[area][k] += v
    for k, v in dist['weekly'].items():
        cwl[area][k] += v


def remove_load(area: str, dist: dict, cl: dict, cwl: dict):
    for k, v in dist['monthly'].items():
        cl[area][k] -= v
    for k, v in dist['weekly'].items():
        cwl[area][k] -= v


def can_assign(area: str, dist: dict, cl: dict, tgt: dict) -> bool:
    """블록 1개 배정 시 월별 목표 초과 여부 확인. tgt에 없으면 무조건 True."""
    if area not in tgt:
        return True
    for (yr, mo), blk_load in dist['monthly'].items():
        if cl.get(area, {}).get((yr, mo), 0.0) + blk_load > tgt[area].get((yr, mo), 0):
            return False
    return True


def can_assign_group(area: str, indices: list, load_dists: list,
                     cl: dict, tgt: dict) -> bool:
    """그룹 전체 배정 시 월별 목표 초과 여부 확인."""
    if area not in tgt:
        return True
    grp_load: dict = defaultdict(float)
    for idx in indices:
        for k, v in load_dists[idx]['monthly'].items():
            grp_load[k] += v
    for (yr, mo), g_load in grp_load.items():
        if cl.get(area, {}).get((yr, mo), 0.0) + g_load > tgt[area].get((yr, mo), 0):
            return False
    return True


def assign_block(df: pd.DataFrame, idx: int, area: str,
                 load_dists: list, cl: dict, cwl: dict):
    df.at[idx, 'assigned_area'] = area
    df.at[idx, 'rule_assigned']  = _current_rule
    add_load(area, load_dists[idx], cl, cwl)


def assign_roundrobin(candidates: list, areas: list,
                      df: pd.DataFrame, load_dists: list,
                      cl: dict, cwl: dict, tgt: dict) -> int:
    """
    candidates를 areas 순환(round-robin) 배정.
    특정 area가 현재 블록을 받지 못하면 그 블록만 건너뜀.
    모든 area가 특정 블록을 거부하면 해당 블록을 skip하고 다음 블록 시도.
    area가 연속으로 거부 횟수가 전체 후보 수를 초과하면 중단.
    반환: 배정 건수
    """
    area_idx    = 0
    cnt         = 0
    skip_streak = 0          # 연속 skip 횟수 (모든 area가 거부한 횟수)
    max_skip    = len(areas) # area 전체가 한 번씩 거부하면 streak 리셋
    for idx in candidates:
        assigned = False
        for offset in range(len(areas)):
            area = areas[(area_idx + offset) % len(areas)]
            if can_assign(area, load_dists[idx], cl, tgt):
                assign_block(df, idx, area, load_dists, cl, cwl)
                area_idx = (area_idx + offset + 1) % len(areas)
                cnt += 1
                skip_streak = 0
                assigned = True
                break
        if not assigned:
            skip_streak += 1
            if skip_streak >= max_skip * 3:
                break          # 연속 미배정이 지속되면 중단
    return cnt


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# R0. 이관물량(H) 처리가능여부 리포팅
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def r0_report_transfer_capacity(df: pd.DataFrame, load_dists: list, capa: dict):
    print("\n" + "=" * 70)
    print("  [R0] 이관물량(H) 처리가능여부 리포팅")
    print("=" * 70)
    year = 2026
    t_load: dict = defaultdict(float)
    h_load: dict = defaultdict(float)
    for i, row in df.iterrows():
        for (yr, mo), load in load_dists[i]['monthly'].items():
            if yr != year:
                continue
            if row['h_t'] == 'T':
                t_load[mo] += load
            else:
                h_load[mo] += load
    total_capa = {mo: sum(c.get(mo, 0) for a, c in capa.items() if a != AREA_Y9)
                  for mo in range(1, 13)}
    alert = False
    for mo in range(1, 13):
        t, h = t_load.get(mo, 0), h_load.get(mo, 0)
        total = t + h * 1.2
        cap   = total_capa.get(mo, 0)
        if total > cap:
            if not alert:
                print("  ⚠️  처리 초과 월 발생:")
                print(f"  {'월':^7} {'T부하':>10} {'H부하×1.2':>12} {'합계':>10} {'전체능력':>10} {'초과':>10}")
                print("  " + "-" * 63)
                alert = True
            print(f"  {year}-{mo:02d}  {t:>10,.0f} {h*1.2:>12,.0f} {total:>10,.0f} {cap:>10,.0f} {total-cap:>10,.0f}")
    if not alert:
        print("  ✅ 전체 월 처리 가능 (초과 없음)")
    print()


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# R1. 작업장별 월별 배정 목표물량 산출
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def r1_calc_target_loads(df: pd.DataFrame, load_dists: list, capa: dict) -> dict:
    """
    월별 전체부하 / 전체능력 = 목표조업도
    목표조업도 × 작업장능력 = 작업장별 배정 목표물량
    반환: {area: {(year, month): target_mh}}
    """
    year = 2026
    total_load: dict = defaultdict(float)
    for dist in load_dists:
        for (yr, mo), load in dist['monthly'].items():
            if yr == year:
                total_load[(yr, mo)] += load
    total_capa = {(year, mo): sum(c.get(mo, 0) for a, c in capa.items() if a != AREA_Y9)
                  for mo in range(1, 13)}
    util = {k: (total_load[k] / total_capa[k] if total_capa[k] > 0 else 0)
            for k in total_capa}
    tgt: dict = {}
    for area, month_caps in capa.items():
        if area == AREA_Y9:
            continue
        tgt[area] = {(year, mo): month_caps.get(mo, 0) * util.get((year, mo), 0)
                     for mo in range(1, 13)}

    print("\n" + "=" * 90)
    print("  [R1] 작업장별 월별 배정 목표물량 (목표조업도 × 능력)")
    print("=" * 90)
    hdr = f"  {'작업장':^8}" + "".join(f"  {m:>2}월" for m in range(1, 13))
    print(hdr)
    print("  " + "-" * (len(hdr) - 2))
    for area in sorted(tgt):
        vals = [tgt[area].get((year, m), 0) for m in range(1, 13)]
        print(f"  {area:^8}" + "".join(f" {v:>6,.0f}" for v in vals))
    util_pct = [util.get((year, m), 0) * 100 for m in range(1, 13)]
    print(f"  {'목표조업도':^8}" + "".join(f" {u:>5.1f}%" for u in util_pct))
    print()
    return tgt


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# R2. H/Y9 블록 → AREA15
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def r2_assign_y9(df: pd.DataFrame, load_dists: list, cl: dict, cwl: dict):
    set_rule('R2')
    mask = (df['h_t'] == 'H') & df['blk'].str.startswith('Y9') & df['assigned_area'].isna()
    cnt = 0
    for idx in df[mask].index:
        assign_block(df, idx, AREA_Y9, load_dists, cl, cwl)
        cnt += 1
    print(f"  [R2] H/Y9 → AREA15: {cnt}건")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# R3. T-중/대 → m_area
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def r3_assign_t_mid_large(df: pd.DataFrame, load_dists: list, cl: dict, cwl: dict):
    set_rule('R3')
    mask = (df['h_t'] == 'T') & df['stg'].isin(['중', '대']) & df['assigned_area'].isna()
    cnt = 0
    for idx in df[mask].index:
        area = df.at[idx, 'm_area']
        if pd.notna(area) and area:
            assign_block(df, idx, area, load_dists, cl, cwl)
            cnt += 1
    print(f"  [R3] T-중/대 → m_area: {cnt}건")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# R4. T-소/area_prior_1=동일 → prt_area
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def r4_assign_t_small_dongil(df: pd.DataFrame, load_dists: list, cl: dict, cwl: dict):
    set_rule('R4')
    mask = (
        (df['h_t'] == 'T') & (df['stg'] == '소') &
        (df['area_prior_1'] == '동일') & df['assigned_area'].isna()
    )
    cnt = 0
    for idx in df[mask].index:
        area = df.at[idx, 'prt_area']
        if pd.notna(area) and area:
            assign_block(df, idx, area, load_dists, cl, cwl)
            cnt += 1
    print(f"  [R4] T-소/동일 → prt_area: {cnt}건")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# R5. H-소 AREA3~5 처리가능여부 리포팅
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def r5_report_h_small(df: pd.DataFrame, load_dists: list, cl: dict, capa: dict):
    print("\n" + "=" * 70)
    print("  [R5] H-소 물량 AREA3~5 처리가능여부 리포팅")
    print("=" * 70)
    year = 2026
    h_so_load: dict = defaultdict(float)
    mask = (df['h_t'] == 'H') & (df['stg'] == '소') & df['assigned_area'].isna()
    for i in df[mask].index:
        for (yr, mo), load in load_dists[i]['monthly'].items():
            if yr == year:
                h_so_load[mo] += load
    area35_remain: dict = defaultdict(float)
    for area in AREAS_SMALL:
        for mo in range(1, 13):
            cap = capa.get(area, {}).get(mo, 0)
            cur = cl.get(area, {}).get((year, mo), 0.0)
            area35_remain[mo] += max(0.0, cap - cur)
    alert = False
    for mo in range(1, 13):
        h   = h_so_load.get(mo, 0)
        rem = area35_remain.get(mo, 0)
        if h > rem:
            if not alert:
                print("  ⚠️  초과 월:")
                print(f"  {'월':^7} {'H-소부하':>12} {'AREA3~5잔여':>14} {'초과':>10}")
                print("  " + "-" * 47)
                alert = True
            print(f"  {year}-{mo:02d}  {h:>12,.0f} {rem:>14,.0f} {h-rem:>10,.0f}")
    if not alert:
        print("  ✅ AREA3~5에서 H-소 전량 처리 가능")
    print()


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# R6. H-소 → AREA3~5 순환(round-robin)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def r6_assign_h_small(df: pd.DataFrame, load_dists: list,
                      cl: dict, cwl: dict, tgt: dict):
    set_rule('R6')
    mask  = (df['h_t'] == 'H') & (df['stg'] == '소') & df['assigned_area'].isna()
    cands = df[mask].sort_values('m_stdt').index.tolist()
    cnt   = assign_roundrobin(cands, AREAS_SMALL, df, load_dists, cl, cwl, tgt)
    print(f"  [R6] H-소 → AREA3~5 순환: {cnt}건")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# R7. H-중-jig_F → AREA2
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def r7_assign_h_mid_jig_f(df: pd.DataFrame, load_dists: list,
                           cl: dict, cwl: dict, tgt: dict):
    set_rule('R7')
    area = 'AREA2'
    SPECIAL_BLK = {'D114', 'D174', 'D204'}

    # Part1: ship_kind=A + 특정 blk 앞4자리 → 목표물량 무관 무조건 AREA2
    mask_p1 = (
        (df['h_t'] == 'H') & (df['ship_kind'] == 'A') & (df['stg'] == '중') &
        df['blk'].str[:4].isin(SPECIAL_BLK) & df['assigned_area'].isna()
    )
    cnt_p1 = 0
    for idx in df[mask_p1].index:
        assign_block(df, idx, area, load_dists, cl, cwl)
        cnt_p1 += 1

    # Part2: H-중/jig=F → AREA2 목표물량까지 (이미 배정된 블록 제외)
    # 특정 월이 목표에 도달해 거부된 블록은 건너뛰고 계속 시도
    mask_p2 = (
        (df['h_t'] == 'H') & (df['stg'] == '중') &
        (df['jig'] == 'F') & df['assigned_area'].isna()
    )
    cnt_p2 = 0
    skip_streak = 0
    for idx in df[mask_p2].sort_values('m_stdt').index:
        if can_assign(area, load_dists[idx], cl, tgt):
            assign_block(df, idx, area, load_dists, cl, cwl)
            cnt_p2 += 1
            skip_streak = 0
        else:
            skip_streak += 1
            if skip_streak >= 50:   # 연속 50건 거부 시 AREA2 용량 소진으로 판단
                break
    print(f"  [R7] H-중-jig_F → AREA2: Part1={cnt_p1}건 + Part2={cnt_p2}건")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# R8. H-중-jig_D/L → AREA1→2→6 순환
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def r8_assign_h_mid_jig_dl(df: pd.DataFrame, load_dists: list,
                            cl: dict, cwl: dict, tgt: dict):
    set_rule('R8')
    mask = (
        (df['h_t'] == 'H') & (df['stg'] == '중') &
        ~df['blk'].str.startswith('H') &
        ~df['blk'].str[:3].isin(['E11', 'E51']) &
        df['jig'].isin(['D', 'L', 'W']) & df['assigned_area'].isna()
    )
    cands = df[mask].sort_values('m_stdt').index.tolist()
    cnt   = assign_roundrobin(cands, AREAS_MID_DL, df, load_dists, cl, cwl, tgt)
    print(f"  [R8] H-중-jig_D/L/W → AREA1→2→6 순환: {cnt}건")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# R9. H-전문화 family 배정
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def r9_assign_specialized_family(df: pd.DataFrame, load_dists: list,
                                  cl: dict, cwl: dict, tgt: dict):
    set_rule('R9')
    def _assign(blk_filter, suffix: str, area: str, label: str):
        base = (df['h_t'] == 'H') & df['stg'].isin(['대', '중']) & df['assigned_area'].isna()
        sub  = df[base & blk_filter & (df['blk'].str[-1] == suffix)].copy()
        sub['_grp'] = sub['pjt'].astype(str) + '|' + sub['blk'].str[:3]
        # 그룹 착수일 오름차순
        grp_order = sub.groupby('_grp')['m_stdt'].min().sort_values().index.tolist()
        cnt = 0
        skip_streak = 0
        for grp in grp_order:
            g_idx = sub[sub['_grp'] == grp].index.tolist()
            for idx in df.loc[g_idx].sort_values('m_stdt').index:
                if df.at[idx, 'assigned_area'] is not None:
                    continue
                if can_assign(area, load_dists[idx], cl, tgt):
                    assign_block(df, idx, area, load_dists, cl, cwl)
                    cnt += 1
                    skip_streak = 0
                else:
                    skip_streak += 1
                    if skip_streak >= 50:
                        print(f"  [R9-{label}] → {area}: {cnt}건")
                        return
        print(f"  [R9-{label}] → {area}: {cnt}건")

    # 1-1: blk 첫째자리 H, 마지막자리 P → AREA7
    _assign(df['blk'].str.startswith('H'), 'P', 'AREA7',  '1-1')
    # 1-2: blk 첫째자리 H, 마지막자리 S → AREA8
    _assign(df['blk'].str.startswith('H'), 'S', 'AREA8',  '1-2')
    # 2-1: blk E11 또는 F51 시작, 마지막자리 P → AREA9
    _assign(df['blk'].str[:3].isin(['E11', 'F51']), 'P', 'AREA9',  '2-1')
    # 2-2: blk E11 또는 F51 시작, 마지막자리 S → AREA10
    _assign(df['blk'].str[:3].isin(['E11', 'F51']), 'S', 'AREA10', '2-2')


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# R10. H-대-jig_D/L 그룹 → AREA12
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def r10_assign_h_large_jig_dl(df: pd.DataFrame, load_dists: list,
                               cl: dict, cwl: dict, tgt: dict):
    set_rule('R10')
    area = 'AREA12'
    seed_mask = (
        (df['h_t'] == 'H') & (df['stg'] == '대') &
        df['jig'].isin(['D', 'L']) & df['assigned_area'].isna()
    )
    # 씨드 블록 그룹키 수집: (pjt, blk[0:3], blk[-1])
    grp_keys = set(
        df[seed_mask].apply(lambda r: (r['pjt'], r['blk'][:3], r['blk'][-1]), axis=1)
    )

    def in_grp(r):
        return (r['pjt'], r['blk'][:3], r['blk'][-1]) in grp_keys

    all_h_unassigned = (df['h_t'] == 'H') & df['assigned_area'].isna()
    cands = (df[all_h_unassigned][df[all_h_unassigned].apply(in_grp, axis=1)]
             .sort_values('m_stdt').index.tolist())
    cnt = 0
    skip_streak = 0
    for idx in cands:
        if can_assign(area, load_dists[idx], cl, tgt):
            assign_block(df, idx, area, load_dists, cl, cwl)
            cnt += 1
            skip_streak = 0
        else:
            skip_streak += 1
            if skip_streak >= 50:
                break
    print(f"  [R10] H-대-jig_D/L 그룹 → AREA12: {cnt}건")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# R11. H-대/중 잔여물량 배정
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def r11_assign_h_remaining(df: pd.DataFrame, load_dists: list,
                            cl: dict, cwl: dict, tgt: dict):
    set_rule('R11')
    mask = (df['h_t'] == 'H') & df['stg'].isin(['대', '중']) & df['assigned_area'].isna()
    sub  = df[mask].copy()
    sub['_grp_key'] = sub['pjt'].astype(str) + '|' + sub['blk'].str[:3] + '|' + sub['blk'].str[-1]
    sub['_ps_key']  = sub['pjt'].astype(str) + '|' + sub['blk'].str[:3]

    # PS 페어 감지: 동일 (pjt, blk[0:3])에 P와 S 접미사 그룹 모두 존재
    sfx_by_ps = sub.groupby('_ps_key')['_grp_key'].apply(lambda ks: {k.split('|')[2] for k in ks})
    ps_pairs  = {k for k, s in sfx_by_ps.items() if 'P' in s and 'S' in s}

    # 그룹 착수일 기준 처리 순서
    grp_min_stdt = sub.groupby('_grp_key')['m_stdt'].min().sort_values()
    processed: set = set()
    cnt = 0

    def assign_individually(idx_list: list):
        nonlocal cnt
        for i in sorted(idx_list, key=lambda x: df.at[x, 'm_stdt']):
            if df.at[i, 'assigned_area'] is not None:
                continue
            for a in AREAS_REMAIN:
                if can_assign(a, load_dists[i], cl, tgt):
                    assign_block(df, i, a, load_dists, cl, cwl)
                    cnt += 1
                    break

    for grp_key in grp_min_stdt.index:
        if grp_key in processed:
            continue
        parts = grp_key.split('|')
        pjt_v, blk3, sfx = int(parts[0]), parts[1], parts[2]
        ps_key    = f"{pjt_v}|{blk3}"
        unassigned = [i for i in sub[sub['_grp_key'] == grp_key].index
                      if df.at[i, 'assigned_area'] is None]
        if not unassigned:
            processed.add(grp_key)
            continue

        # ── PS 페어 배정 ─────────────────────────────────────────────────────
        if ps_key in ps_pairs and sfx in ('P', 'S'):
            partner_sfx = 'S' if sfx == 'P' else 'P'
            partner_key = f"{pjt_v}|{blk3}|{partner_sfx}"
            processed.add(grp_key)
            processed.add(partner_key)

            p_key = grp_key     if sfx == 'P' else partner_key
            s_key = partner_key if sfx == 'P' else grp_key
            p_idx = [i for i in sub[sub['_grp_key'] == p_key].index
                     if df.at[i, 'assigned_area'] is None]
            s_idx = [i for i in sub[sub['_grp_key'] == s_key].index
                     if df.at[i, 'assigned_area'] is None]

            pair_done = False
            for (p_area, s_area) in PS_AREA_PAIRS:
                if (can_assign_group(p_area, p_idx, load_dists, cl, tgt) and
                        can_assign_group(s_area, s_idx, load_dists, cl, tgt)):
                    for i in p_idx:
                        assign_block(df, i, p_area, load_dists, cl, cwl)
                        cnt += 1
                    for i in s_idx:
                        assign_block(df, i, s_area, load_dists, cl, cwl)
                        cnt += 1
                    pair_done = True
                    break
            if not pair_done:
                assign_individually(p_idx + s_idx)

        # ── 단독 그룹 배정 ───────────────────────────────────────────────────
        else:
            processed.add(grp_key)
            grp_done = False
            for a in AREAS_REMAIN:
                if can_assign_group(a, unassigned, load_dists, cl, tgt):
                    for i in unassigned:
                        assign_block(df, i, a, load_dists, cl, cwl)
                        cnt += 1
                    grp_done = True
                    break
            if not grp_done:
                assign_individually(unassigned)

    print(f"  [R11] H-대/중 잔여 → AREA7~14: {cnt}건")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# R12. H-dir 블록 → prt_area (override)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def r12_assign_h_dir(df: pd.DataFrame, load_dists: list, cl: dict, cwl: dict):
    set_rule('R12')
    mask = (df['h_t'] == 'H') & df['blk'].str[3].isin(['9', 'H'])
    cnt = 0
    for idx in df[mask].index:
        area = df.at[idx, 'prt_area']
        if not (pd.notna(area) and area):
            continue  # prt_area 공란 → skip
        old = df.at[idx, 'assigned_area']
        if old and old != area:
            remove_load(old, load_dists[idx], cl, cwl)
        assign_block(df, idx, area, load_dists, cl, cwl)
        cnt += 1
    print(f"  [R12] H-dir(blk[3]=9/H) → prt_area override: {cnt}건")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# R13. T-소 분산배정
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

AREA_PRIOR_COLS = ['area_prior_1', 'area_prior_2', 'area_prior_3',
                   'area_prior_4', 'area_prior_5']

def r13_assign_t_small_dist(df: pd.DataFrame, load_dists: list,
                             cl: dict, cwl: dict, tgt: dict,
                             area_seq: list):
    """
    T-소 미배정 블록을 area_seq 순서로 분산 배정.
    - (pjt, blk, type) 그룹핑
    - 그룹 정렬: assign_prior 오름차순(낮은수=높은우선순위) → min(m_stdt) 오름차순
    - area별: area_prior_1 매칭 우선 → 미달 시 area_prior_2~5 순차 추가
    - 목표물량 도달 시 해당 area 중단, 다음 area 진행
    """
    set_rule('R13')
    total_cnt = 0

    for area in area_seq:
        if area not in tgt:
            continue
        area_cnt = 0

        for prior_col in AREA_PRIOR_COLS:
            # 현재 prior_col에서 area와 매칭되는 미배정 T-소 블록
            sub_mask = (
                (df['h_t'] == 'T') & (df['stg'] == '소') &
                df['assigned_area'].isna() & (df[prior_col] == area)
            )
            sub = df[sub_mask].copy()
            if sub.empty:
                continue

            # (pjt, blk, type) 그룹키
            sub['_grp'] = (sub['pjt'].astype(str) + '|' +
                           sub['blk'].astype(str) + '|' +
                           sub['type'].astype(str))

            # 그룹별 대표 assign_prior(최솟값)와 최소 m_stdt로 그룹 순서 결정
            grp_info = (sub.groupby('_grp')
                        .agg(ap=('assign_prior', 'min'), ms=('m_stdt', 'min'))
                        .sort_values(['ap', 'ms']))
            grp_order = grp_info.index.tolist()

            area_skip = 0   # 이 prior_col 내 연속 거부 횟수

            for grp_key in grp_order:
                grp_idx = sub[sub['_grp'] == grp_key].sort_values('m_stdt').index.tolist()

                for idx in grp_idx:
                    if df.at[idx, 'assigned_area'] is not None:
                        continue
                    if can_assign(area, load_dists[idx], cl, tgt):
                        assign_block(df, idx, area, load_dists, cl, cwl)
                        area_cnt += 1
                        total_cnt += 1
                        area_skip = 0
                    else:
                        area_skip += 1

                if area_skip >= 50:
                    break   # 이 prior_col 내 연속 거부 50건 → 다음 prior_col 시도

        print(f"    [R13] {area}: {area_cnt}건")

    print(f"  [R13] T-소 분산배정 합계: {total_cnt}건")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 결과 요약 출력 & 저장
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def print_utilization_summary(df: pd.DataFrame, cl: dict, capa: dict):
    year = 2026
    print("\n" + "=" * 90)
    print("  배정 결과 — 작업장별 월별 조업도 (배정부하 / 능력 × 100%)")
    print("=" * 90)
    hdr = f"  {'작업장':^8}" + "".join(f"  {m:>2}월" for m in range(1, 13)) + "  연평균"
    print(hdr)
    print("  " + "-" * (len(hdr) - 2))
    all_utils = []
    for area in sorted(a for a in capa if a != AREA_Y9):
        utils = []
        for m in range(1, 13):
            cap  = capa[area].get(m, 0)
            load = cl.get(area, {}).get((year, m), 0.0)
            utils.append((load / cap * 100) if cap > 0 else 0.0)
        all_utils.append(utils)
        avg = sum(utils) / len(utils)
        print(f"  {area:^8}" + "".join(f" {u:>5.1f}%" for u in utils) + f" {avg:>6.1f}%")
    print("  " + "-" * (len(hdr) - 2))
    col_avg = [sum(r[i] for r in all_utils) / max(len(all_utils), 1) for i in range(12)]
    print(f"  {'전체평균':^8}" + "".join(f" {u:>5.1f}%" for u in col_avg) + f" {sum(col_avg)/12:>6.1f}%")

    total = len(df)
    asgn  = df['assigned_area'].notna().sum()
    print(f"\n  전체: {total:,}건 | 배정완료: {asgn:,}건 | 미배정: {total-asgn:,}건")
    print("\n  [작업장별 배정 건수]")
    for area, c in df['assigned_area'].value_counts().sort_index().items():
        print(f"    {area}: {c:,}건")


RULE_LABELS = {
    'R2':  'H/Y9 → AREA15',
    'R3':  'T-중/대 → m_area',
    'R4':  'T-소/동일 → prt_area',
    'R6':  'H-소 → AREA3~5 순환',
    'R7':  'H-중-jig_F → AREA2',
    'R8':  'H-중-jig_D/L → AREA1→2→6',
    'R9':  'H-전문화 Family',
    'R10': 'H-대-jig_D/L 그룹 → AREA12',
    'R11': 'H-대/중 잔여 → AREA7~14',
    'R12': 'H-dir → prt_area(override)',
    'R13': 'T-소 분산배정 (area_seq 순)',
}

def save_stats(df: pd.DataFrame, cl: dict, capa: dict, tgt: dict):
    year = 2026
    # 단계별 배정 건수 & 작업장 분포
    steps = []
    for rule in ['R2','R3','R4','R6','R7','R8','R9','R10','R11','R12']:
        sub = df[df['rule_assigned'] == rule]
        area_dist = sub['assigned_area'].value_counts().to_dict()
        steps.append({
            'rule':   rule,
            'label':  RULE_LABELS.get(rule, rule),
            'count':  int(len(sub)),
            'areas':  {k: int(v) for k, v in area_dist.items()},
        })

    # 작업장별 월별 조업도
    utilization = {}
    for area in sorted(a for a in capa if a != AREA_Y9):
        utilization[area] = {}
        for mo in range(1, 13):
            cap  = capa[area].get(mo, 0)
            load = cl.get(area, {}).get((year, mo), 0.0)
            util = round(load / cap * 100, 1) if cap > 0 else 0.0
            utilization[area][f'{year}-{mo:02d}'] = util

    # 작업장별 목표 조업도
    target_util = {}
    for area in sorted(a for a in capa if a != AREA_Y9):
        target_util[area] = {}
        for mo in range(1, 13):
            cap = capa[area].get(mo, 0)
            t   = tgt.get(area, {}).get((year, mo), 0)
            target_util[area][f'{year}-{mo:02d}'] = round(t / cap * 100, 1) if cap > 0 else 0.0

    stats = {
        'total':       int(len(df)),
        'assigned':    int(df['assigned_area'].notna().sum()),
        'unassigned':  int(df['assigned_area'].isna().sum()),
        'steps':       steps,
        'utilization': utilization,
        'target_util': target_util,
    }
    with open(STATS_PATH, 'w', encoding='utf-8') as f:
        json.dump(stats, f, ensure_ascii=False, indent=2)
    print(f"📊 통계 저장 완료: {STATS_PATH}")


def save_result(df: pd.DataFrame, path: str):
    out = df.copy()
    out['m_stdt'] = out['m_stdt'].dt.date
    out['m_fndt'] = out['m_fndt'].dt.date
    out.to_excel(path, index=False)
    print(f"\n✅ 저장 완료: {path}")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 메인
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

if __name__ == '__main__':
    print("=" * 70)
    print("  블록 작업장 배정 Agent")
    print("=" * 70)

    print("\n▶ 데이터 로드 ...")
    df       = load_blk_master(BLK_MASTER_PATH)
    capa     = load_area_capa(AREA_CAPA_PATH)
    area_seq = load_area_seq(AREA_CAPA_PATH)
    print(f"  blk_master : {len(df):,}행")
    print(f"  area_capa  : {len(capa)}개 작업장 ({', '.join(sorted(capa.keys()))})")

    print("\n▶ 글로벌 부하 분산 계산 (load_mh → 월별/주별) ...")
    load_dists = [calc_load_distribution(row) for _, row in df.iterrows()]

    # 누적 부하 추적 테이블 (월별 + 주별)
    all_areas = list(capa.keys()) + [AREA_Y9]
    cl:  dict = {a: defaultdict(float) for a in all_areas}
    cwl: dict = {a: defaultdict(float) for a in all_areas}

    print("\n▶ 순차 배정 규칙 실행 ...")
    r0_report_transfer_capacity(df, load_dists, capa)
    tgt = r1_calc_target_loads(df, load_dists, capa)
    r2_assign_y9(df, load_dists, cl, cwl)
    r3_assign_t_mid_large(df, load_dists, cl, cwl)
    r4_assign_t_small_dongil(df, load_dists, cl, cwl)
    r5_report_h_small(df, load_dists, cl, capa)
    r6_assign_h_small(df, load_dists, cl, cwl, tgt)
    r7_assign_h_mid_jig_f(df, load_dists, cl, cwl, tgt)
    r8_assign_h_mid_jig_dl(df, load_dists, cl, cwl, tgt)
    r9_assign_specialized_family(df, load_dists, cl, cwl, tgt)
    r10_assign_h_large_jig_dl(df, load_dists, cl, cwl, tgt)
    r11_assign_h_remaining(df, load_dists, cl, cwl, tgt)
    r12_assign_h_dir(df, load_dists, cl, cwl)
    r13_assign_t_small_dist(df, load_dists, cl, cwl, tgt, area_seq)

    print_utilization_summary(df, cl, capa)
    save_stats(df, cl, capa, tgt)
    save_result(df, RESULT_PATH)
