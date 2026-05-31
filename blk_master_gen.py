"""
blk_master_gen.py — 블록 마스터 Excel 데이터 생성기
3,200행 × 24컬럼 (알파벳 데이터 대문자)
"""

import pandas as pd
import numpy as np
import random
import string
from collections import defaultdict
from datetime import date, timedelta

SEED = 42
np.random.seed(SEED)
random.seed(SEED)

N = 3_200

# ─────────────────────────────────────────────────────────────
# 대한민국 2026년 공휴일
# ─────────────────────────────────────────────────────────────
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
}

def is_workday(d: date) -> bool:
    return d.weekday() < 5 and d not in KR_HOLIDAYS

def add_workdays(start: date, n: int) -> date:
    cur = start
    count = 0
    while count < n:
        cur += timedelta(days=1)
        if is_workday(cur):
            count += 1
    return cur

workdays_by_month: dict = defaultdict(list)
d = date(2026, 1, 1)
while d <= date(2026, 12, 30):
    if is_workday(d):
        workdays_by_month[d.month].append(d)
    d += timedelta(days=1)

AREAS = [f'AREA{i}' for i in range(1, 15)]
TYPE_COMBOS = [x + y for x in 'ABCD' for y in 'ABCD']  # 16가지 AA~DD

# ─────────────────────────────────────────────────────────────
# 1. pjt — 1000~1020, 21개 값, 3200행 골고루
# ─────────────────────────────────────────────────────────────
pjt_values = list(range(1000, 1021))  # 21개
base_cnt = N // len(pjt_values)
remainder = N % len(pjt_values)
pjt_pool = []
for i, v in enumerate(pjt_values):
    pjt_pool.extend([v] * (base_cnt + (1 if i < remainder else 0)))
random.shuffle(pjt_pool)
pjt = np.array(pjt_pool, dtype=int)

# ─────────────────────────────────────────────────────────────
# 2. ship_kind — pjt별 그룹, A~D 랜덤 (동일 pjt → 동일 값)
# ─────────────────────────────────────────────────────────────
pjt_ship_map = {v: random.choice('ABCD') for v in pjt_values}
ship_kind = np.array([pjt_ship_map[p] for p in pjt])

# ─────────────────────────────────────────────────────────────
# 3. h_t — pjt<1016 → H, pjt>1015 → T
# ─────────────────────────────────────────────────────────────
h_t = np.array(['H' if p < 1016 else 'T' for p in pjt])

# ─────────────────────────────────────────────────────────────
# 4. blk — 5자리 코드
#   H: [A-Z][0-9][0-9][0-4][P/S/C]  (P 개수 == S 개수)
#   T: 동일 blk 3~10개 그룹 [A-Z][0-9][0-9][0-4][A/B/D/E]
# ─────────────────────────────────────────────────────────────
blk = np.full(N, '', dtype='U5')

# H 행: P/S/C 접미사, P와 S 개수 동일
h_idx = np.where(h_t == 'H')[0]
n_h = len(h_idx)
n_p = n_h // 3
n_s = n_p
n_c = n_h - n_p - n_s
h_sfx = ['P'] * n_p + ['S'] * n_s + ['C'] * n_c
random.shuffle(h_sfx)
for pos, idx in enumerate(h_idx):
    blk[idx] = (random.choice(string.ascii_uppercase) +
                str(random.randint(0, 9)) +
                str(random.randint(0, 9)) +
                str(random.randint(0, 4)) +
                h_sfx[pos])

# T 행: pjt별로 그룹핑, 동일 blk 3~10개씩 반복
t_idx_by_pjt = defaultdict(list)
for idx in np.where(h_t == 'T')[0]:
    t_idx_by_pjt[pjt[idx]].append(idx)

for pjt_val, indices in t_idx_by_pjt.items():
    pos = 0
    while pos < len(indices):
        g = min(random.randint(3, 10), len(indices) - pos)
        blk_val = (random.choice(string.ascii_uppercase) +
                   str(random.randint(0, 9)) +
                   str(random.randint(0, 9)) +
                   str(random.randint(0, 4)) +
                   random.choice('ABDE'))
        for j in range(g):
            blk[indices[pos + j]] = blk_val
        pos += g

# ─────────────────────────────────────────────────────────────
# 5. stg
#   blk[3]=='0' → 대
#   T 그룹 첫 행 (blk[3]!='0') → 중  (prt_area 참조 보장)
#   H 첫 행 (blk[3]!='0')      → 중/소 랜덤
#   같은 (pjt, blk) 중복 → 소
# ─────────────────────────────────────────────────────────────
stg = np.empty(N, dtype='U2')
seen_pb: dict = {}
for i in range(N):
    key = (int(pjt[i]), blk[i])
    if key in seen_pb:
        stg[i] = '소'
    else:
        seen_pb[key] = i
        if blk[i][3] == '0':
            stg[i] = '대'
        elif h_t[i] == 'T':
            stg[i] = '중'
        else:
            stg[i] = random.choice(['중', '소'])

# ─────────────────────────────────────────────────────────────
# 6. sub_blk — h_t=T AND stg=소: [A-Z][A-Z][01~20] (4자리)
# ─────────────────────────────────────────────────────────────
sub_blk = np.full(N, '', dtype='U4')
for idx in np.where((h_t == 'T') & (stg == '소'))[0]:
    sub_blk[idx] = (random.choice(string.ascii_uppercase) +
                    random.choice(string.ascii_uppercase) +
                    f'{random.randint(1, 20):02d}')

# ─────────────────────────────────────────────────────────────
# 7. type — h_t=T: 16가지 조합(AA~DD) 랜덤
# ─────────────────────────────────────────────────────────────
type_col = np.full(N, '', dtype='U2')
for idx in np.where(h_t == 'T')[0]:
    type_col[idx] = random.choice(TYPE_COMBOS)

# ─────────────────────────────────────────────────────────────
# 8. jig — h_t=H: D/F/L/W 랜덤
# ─────────────────────────────────────────────────────────────
jig = np.full(N, '', dtype='U1')
for idx in np.where(h_t == 'H')[0]:
    jig[idx] = random.choice('DFLW')

# ─────────────────────────────────────────────────────────────
# 9. e_wt — stg=중/대: 15~200, stg=소: 빈값
# ─────────────────────────────────────────────────────────────
e_wt = np.full(N, np.nan)
for idx in np.where((stg == '중') | (stg == '대'))[0]:
    e_wt[idx] = float(random.randint(15, 200))

# ─────────────────────────────────────────────────────────────
# 10. wt — T/소: 1~10, 그 외: 3~60 (e_wt 존재 시 ≤ e_wt)
# ─────────────────────────────────────────────────────────────
wt = np.full(N, np.nan)
for i in range(N):
    if h_t[i] == 'T' and stg[i] == '소':
        wt[i] = float(random.randint(1, 10))
    elif not np.isnan(e_wt[i]):
        upper = min(60, int(e_wt[i]))
        wt[i] = float(random.randint(3, upper))
    else:
        wt[i] = float(random.randint(3, 60))

# ─────────────────────────────────────────────────────────────
# 11. w_len — H/중: 50~300, H/대: 300~1000
# ─────────────────────────────────────────────────────────────
w_len = np.full(N, np.nan)
for idx in np.where((h_t == 'H') & (stg == '중'))[0]:
    w_len[idx] = float(random.randint(50, 300))
for idx in np.where((h_t == 'H') & (stg == '대'))[0]:
    w_len[idx] = float(random.randint(300, 1000))

# ─────────────────────────────────────────────────────────────
# 12. sub_w_len — H/소: 50%는 빈값, 50%는 200~1500
# ─────────────────────────────────────────────────────────────
sub_w_len = np.full(N, np.nan)
h_so_all = list(np.where((h_t == 'H') & (stg == '소'))[0])
fill_idx = random.sample(h_so_all, len(h_so_all) // 2)
for idx in fill_idx:
    sub_w_len[idx] = float(random.randint(200, 1500))

# ─────────────────────────────────────────────────────────────
# 13. m_stdt — 2026 영업일, 월별 균등 분포
# ─────────────────────────────────────────────────────────────
rows_per_month = N // 12
m_stdt_list = []
for month in range(1, 13):
    cnt = rows_per_month + (1 if month <= N % 12 else 0)
    m_stdt_list.extend(random.choices(workdays_by_month[month], k=cnt))
random.shuffle(m_stdt_list)
m_stdt = np.array(m_stdt_list, dtype=object)

# ─────────────────────────────────────────────────────────────
# 14. m_fndt — m_stdt + 영업일 추가
# ─────────────────────────────────────────────────────────────
m_fndt = np.empty(N, dtype=object)
for i in range(N):
    ht, sg = h_t[i], stg[i]
    if   ht == 'H' and sg == '소': nd = random.randint(3, 5)
    elif ht == 'H':                 nd = random.randint(10, 18)
    elif ht == 'T' and sg == '소': nd = random.randint(4, 30)
    else:                           nd = random.randint(10, 55)
    m_fndt[i] = add_workdays(m_stdt[i], nd)

# ─────────────────────────────────────────────────────────────
# 15. m_area — AREA1~14 랜덤
# ─────────────────────────────────────────────────────────────
m_area = np.array([random.choice(AREAS) for _ in range(N)])

# ─────────────────────────────────────────────────────────────
# 16. m_mh — T행만: 대 500~2000, 중 2~200, 소 1~100
# ─────────────────────────────────────────────────────────────
m_mh = np.full(N, np.nan)
for i in range(N):
    if h_t[i] == 'T':
        sg = stg[i]
        if   sg == '대': m_mh[i] = float(random.randint(500, 2000))
        elif sg == '중': m_mh[i] = float(random.randint(2, 200))
        else:            m_mh[i] = float(random.randint(1, 100))

# ─────────────────────────────────────────────────────────────
# 17. prt_area — T/소 행: 같은 (pjt, blk)의 중/대 행 m_area 참조
# ─────────────────────────────────────────────────────────────
parent_m_area: dict = {}
for i in range(N):
    if stg[i] in ('중', '대'):
        key = (int(pjt[i]), blk[i])
        parent_m_area.setdefault(key, m_area[i])

prt_area = np.full(N, '', dtype='U10')
for i in range(N):
    if h_t[i] == 'T' and stg[i] == '소':
        key = (int(pjt[i]), blk[i])
        prt_area[i] = parent_m_area.get(key, '')

# ─────────────────────────────────────────────────────────────
# 18. load_mh
#   T       → m_mh
#   H/중·대 → w_len / 0.88
#   H/소    → sub_w_len / 3.93
# ─────────────────────────────────────────────────────────────
load_mh = np.full(N, np.nan)
for i in range(N):
    if h_t[i] == 'T':
        load_mh[i] = m_mh[i]
    elif stg[i] in ('중', '대'):
        if not np.isnan(w_len[i]):
            load_mh[i] = w_len[i] / 0.88
    else:  # H/소
        if not np.isnan(sub_w_len[i]):
            load_mh[i] = sub_w_len[i] / 3.93

# ─────────────────────────────────────────────────────────────
# 19. area_prior_1 — T/소 10% → '동일', 나머지 → AREA 랜덤
# ─────────────────────────────────────────────────────────────
ts_so_idx = list(np.where((h_t == 'T') & (stg == '소'))[0])
dongil_set = set(random.sample(ts_so_idx, max(1, len(ts_so_idx) // 10)))
area_prior_1 = np.array([
    '동일' if i in dongil_set else random.choice(AREAS)
    for i in range(N)
], dtype='U10')

# ─────────────────────────────────────────────────────────────
# 20~23. area_prior_2~5 — 이전 컬럼 제외한 AREA 랜덤
# ─────────────────────────────────────────────────────────────
def gen_next_prior(*prev_cols):
    result = np.empty(N, dtype='U10')
    for i in range(N):
        excluded = set()
        for c in prev_cols:
            v = c[i]
            if v == '동일':
                excluded.add(m_area[i])  # '동일'은 해당 행의 m_area로 해석
            elif v:
                excluded.add(v)
        pool = [a for a in AREAS if a not in excluded]
        result[i] = random.choice(pool)
    return result

area_prior_2 = gen_next_prior(area_prior_1)
area_prior_3 = gen_next_prior(area_prior_1, area_prior_2)
area_prior_4 = gen_next_prior(area_prior_1, area_prior_2, area_prior_3)
area_prior_5 = gen_next_prior(area_prior_1, area_prior_2, area_prior_3, area_prior_4)

# ─────────────────────────────────────────────────────────────
# 24. assign_prior — type 그룹별 1~16
# ─────────────────────────────────────────────────────────────
type_to_num = {t: i + 1 for i, t in enumerate(TYPE_COMBOS)}
assign_prior = pd.array(
    [type_to_num[t] if t else pd.NA for t in type_col],
    dtype='Int64'
)

# ─────────────────────────────────────────────────────────────
# DataFrame 조립
# ─────────────────────────────────────────────────────────────
df = pd.DataFrame({
    'pjt':          pjt,
    'ship_kind':    ship_kind,
    'h_t':          h_t,
    'blk':          blk,
    'stg':          stg,
    'sub_blk':      sub_blk,
    'type':         type_col,
    'jig':          jig,
    'e_wt':         e_wt,
    'wt':           wt,
    'w_len':        w_len,
    'sub_w_len':    sub_w_len,
    'm_stdt':       m_stdt,
    'm_fndt':       m_fndt,
    'm_area':       m_area,
    'm_mh':         m_mh,
    'prt_area':     prt_area,
    'load_mh':      load_mh,
    'area_prior_1': area_prior_1,
    'area_prior_2': area_prior_2,
    'area_prior_3': area_prior_3,
    'area_prior_4': area_prior_4,
    'area_prior_5': area_prior_5,
    'assign_prior': assign_prior,
})

df['m_stdt'] = pd.to_datetime(df['m_stdt'])
df['m_fndt'] = pd.to_datetime(df['m_fndt'])

# ─────────────────────────────────────────────────────────────
# 검증 출력
# ─────────────────────────────────────────────────────────────
print("=" * 60)
print(f"행 수: {len(df):,} | 열 수: {len(df.columns)}")
print()
print("[h_t]", df['h_t'].value_counts().to_dict())
print("[stg]", df['stg'].value_counts().to_dict())
print("[ship_kind]", df['ship_kind'].value_counts().sort_index().to_dict())
print("[jig (H행)]", df[df['h_t']=='H']['jig'].value_counts().sort_index().to_dict())
h_blk = df[df['h_t']=='H']['blk']
p_cnt = (h_blk.str[-1] == 'P').sum()
s_cnt = (h_blk.str[-1] == 'S').sum()
print(f"[blk P/S 균등] P={p_cnt}, S={s_cnt} → {'✓' if p_cnt == s_cnt else '✗'}")
print(f"[type 조합 수] {df[df['h_t']=='T']['type'].nunique()}가지")
print(f"[sub_blk 채워진 행] {(df['sub_blk']!='').sum():,}")
print(f"[e_wt 채워진 행] {df['e_wt'].notna().sum():,}")
print(f"[wt <= e_wt 위반] {((df['wt'] > df['e_wt']) & df['e_wt'].notna()).sum()}")
print(f"[w_len 채워진 행] {df['w_len'].notna().sum():,}")
print(f"[sub_w_len 채워진 행] {df['sub_w_len'].notna().sum():,}")
print(f"[m_mh 채워진 행] {df['m_mh'].notna().sum():,}")
print(f"[prt_area 채워진 행] {(df['prt_area']!='').sum():,}")
print(f"[load_mh 채워진 행] {df['load_mh'].notna().sum():,}")
print(f"[area_prior_1 동일 개수] {(df['area_prior_1']=='동일').sum()}")
print(f"[assign_prior 범위] {df['assign_prior'].dropna().min()}~{df['assign_prior'].dropna().max()}")
print()
print("[m_stdt 월별 분포]")
print(df['m_stdt'].dt.month.value_counts().sort_index().to_string())

# ─────────────────────────────────────────────────────────────
# Excel 저장
# ─────────────────────────────────────────────────────────────
output_path = '/mnt/d/mook/AI/pjt/company/blk_assign/blk_master.xlsx'
df.to_excel(output_path, index=False)
print(f"\n✅ 저장 완료: {output_path}")
