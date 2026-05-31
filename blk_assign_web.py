"""
blk_assign_web.py — 블록 배정 결과 HTML 대시보드 생성기
생성: blk_assign_report.html (브라우저에서 직접 열기)

사용법:
  python3 blk_assign_web.py
  → blk_assign_report.html 생성 → 브라우저로 열기
"""

import json
import pandas as pd
from pathlib import Path

BASE_DIR    = '/mnt/d/mook/AI/pjt/company/blk_assign'
RESULT_PATH = f'{BASE_DIR}/blk_assign_result.xlsx'
STATS_PATH  = f'{BASE_DIR}/blk_assign_stats.json'
OUTPUT_PATH = f'{BASE_DIR}/blk_assign_report.html'

# ── 데이터 로드 ───────────────────────────────────────────────────────────────
print("▶ 데이터 로드 중...")
df    = pd.read_excel(RESULT_PATH)
stats = json.loads(Path(STATS_PATH).read_text(encoding='utf-8'))

df['m_stdt'] = pd.to_datetime(df['m_stdt']).dt.strftime('%Y-%m-%d')
df['m_fndt'] = pd.to_datetime(df['m_fndt']).dt.strftime('%Y-%m-%d')
df['assigned_area']  = df['assigned_area'].fillna('')
df['rule_assigned']  = df['rule_assigned'].fillna('')
df['prt_area']       = df['prt_area'].fillna('')

# ── 블록 목록 JSON (필요 컬럼만) ──────────────────────────────────────────────
block_cols = ['pjt', 'h_t', 'blk', 'stg', 'm_stdt', 'm_fndt',
              'm_area', 'load_mh', 'assigned_area', 'rule_assigned']
blocks_json = df[block_cols].fillna('').to_dict(orient='records')

# ── Mermaid 플로우챠트 정의 ───────────────────────────────────────────────────
MERMAID_CHART = """
flowchart TD
    START([▶ 시작]) --> LOAD[📂 데이터 로드<br/>blk_master + area_capa]
    LOAD --> CALC[⚙️ 글로벌: load_mh 부하 분산<br/>m_stdt~m_fndt 영업일 균등 배분<br/>→ 월별 부하 + 주별 부하 집계]

    CALC --> R0{📊 R0 이관물량 처리가능 리포팅<br/>T부하 + H부하×1.2 vs 전체능력<br/>초과 시 월별 경고 출력}
    R0 --> R1[📌 R1 목표물량 산출<br/>목표조업도 = 전체부하 ÷ 전체능력<br/>배정목표 = 목표조업도 × 작업장능력]

    R1 --> R2[R2 H행 + blk Y9시작<br/>→ AREA15 무조건 배정]
    R2 --> R3[R3 T행 + stg=중·대<br/>→ m_area 배정]
    R3 --> R4[R4 T행 + stg=소 + area_prior_1=동일<br/>→ prt_area 배정]

    R4 --> R5{📊 R5 H-소 처리가능 리포팅<br/>H-소 부하 vs AREA3~5 잔여능력<br/>초과 시 월별 경고 출력}
    R5 --> R6[R6 H행 + stg=소<br/>AREA3 → AREA4 → AREA5 순환 배정<br/>목표물량 도달 시 해당 area 중단]

    R6 --> R7[R7-Part1 H+중+ship_kind=A+blk∈D114·D174·D204<br/>→ AREA2 무조건 배정<br/>R7-Part2 H+중+jig=F<br/>→ AREA2 목표물량까지 배정]

    R7 --> R8[R8 H행 + stg=중 + jig=D·L<br/>단 H·E11·E51 시작 블록 제외<br/>AREA1 → AREA2 → AREA6 순환 배정]

    R8 --> R9[R9 H행 + stg=대·중 전문화 Family<br/>1-1: blk H시작 + 말미P → AREA7<br/>1-2: blk H시작 + 말미S → AREA8<br/>2-1: blk E11·F51시작 + 말미P → AREA9<br/>2-2: blk E11·F51시작 + 말미S → AREA10]

    R9 --> R10[R10 H행 + stg=대 + jig=D·L<br/>동일 그룹 포함하여 → AREA12<br/>목표물량 도달 시 중단]

    R10 --> R11[R11 H행 + stg=대·중 잔여물량<br/>PS페어 그룹: AREA7+8 · AREA9+10 · AREA13+14<br/>단독 그룹: AREA7→14 순차 first-fit<br/>그룹 초과 시 해제 후 개별 배정]

    R11 --> R12[R12 H행 + blk 4번째 자리 9 또는 H<br/>→ prt_area로 override 배정]
    R12 --> SAVE[💾 blk_assign_result.xlsx 저장]

    style R0 fill:#fff3cd,stroke:#ffc107
    style R5 fill:#fff3cd,stroke:#ffc107
    style CALC fill:#d1ecf1,stroke:#0c7cd5
    style R1 fill:#d4edda,stroke:#28a745
    style SAVE fill:#d4edda,stroke:#28a745
"""

# ── 조업도 색상 (0~150%) ──────────────────────────────────────────────────────
def util_color(v: float) -> str:
    if v == 0:
        return '#f8f9fa'
    elif v < 60:
        return f'hsl(210,70%,{max(75, 95 - v * 0.3):.0f}%)'
    elif v < 90:
        return f'hsl(120,60%,{max(55, 85 - (v-60)*0.7):.0f}%)'
    elif v < 110:
        return f'hsl(45,90%,{max(60, 80 - (v-90)*.5):.0f}%)'
    else:
        return f'hsl(0,80%,{max(50, 75 - (v-110)*.4):.0f}%)'

# ── 조업도 HTML 테이블 생성 ────────────────────────────────────────────────────
months = [f'2026-{m:02d}' for m in range(1, 13)]
util   = stats['utilization']
tgt_u  = stats['target_util']

util_rows = ''
for area in sorted(util.keys()):
    vals = [util[area].get(m, 0) for m in months]
    avg  = sum(vals) / len(vals)
    tgt_vals = [tgt_u.get(area, {}).get(m, 0) for m in months]
    cells = ''
    for i, v in enumerate(vals):
        bg = util_color(v)
        tv = tgt_vals[i]
        cells += f'<td style="background:{bg};text-align:center;font-size:12px" title="목표:{tv:.1f}%">{v:.1f}%</td>'
    avg_bg = util_color(avg)
    util_rows += f'<tr><td class="fw-semibold">{area}</td>{cells}<td style="background:{avg_bg};text-align:center;font-size:12px">{avg:.1f}%</td></tr>\n'

# ── 단계별 결과 테이블 ─────────────────────────────────────────────────────────
step_rows = ''
cumulative = 0
for s in stats['steps']:
    cumulative += s['count']
    areas_str = ', '.join(f"{a}:{c}" for a, c in sorted(s['areas'].items()))
    bar_width = min(100, s['count'] * 2) if s['count'] > 0 else 0
    step_rows += f"""<tr>
        <td><span class="badge bg-primary">{s['rule']}</span></td>
        <td>{s['label']}</td>
        <td class="text-end fw-bold">{s['count']:,}</td>
        <td class="text-end text-muted">{cumulative:,}</td>
        <td><div style="background:#0d6efd;height:16px;width:{bar_width}px;border-radius:3px;display:inline-block"></div></td>
        <td style="font-size:11px">{areas_str or '-'}</td>
    </tr>\n"""

# ── 전체 HTML ─────────────────────────────────────────────────────────────────
MONTHS_HDR = ''.join(f'<th>{m[5:]}</th>' for m in months)

HTML = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>블록 작업장 배정 Dashboard</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script type="module">
  import mermaid from 'https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.esm.min.mjs';
  mermaid.initialize({{ startOnLoad:true, theme:'default', flowchart:{{curve:'basis'}} }});
</script>
<style>
body{{ font-family:'Segoe UI',sans-serif; background:#f5f7fa; }}
.nav-tabs .nav-link{{ color:#495057; font-weight:500; }}
.nav-tabs .nav-link.active{{ color:#0d6efd; border-bottom:3px solid #0d6efd; }}
.card{{ border:none; border-radius:12px; box-shadow:0 2px 8px rgba(0,0,0,.08); }}
.stat-box{{ text-align:center; padding:20px; border-radius:10px; }}
.mermaid{{ max-width:100%; overflow-x:auto; }}
#blockTable{{ font-size:13px; }}
#blockTable td,#blockTable th{{ padding:5px 8px; }}
.filter-bar {{ background:#fff; padding:12px 16px; border-radius:8px; margin-bottom:12px; }}
.legend-item{{ display:inline-block; width:16px; height:16px; border-radius:3px; margin-right:4px; vertical-align:middle; }}
</style>
</head>
<body>

<nav class="navbar navbar-dark bg-dark px-4 py-2">
  <span class="navbar-brand fw-bold">🏗️ 블록 작업장 배정 Dashboard</span>
  <span class="text-light small">전체: {stats['total']:,}건 | 배정완료: {stats['assigned']:,}건 | 미배정: {stats['unassigned']:,}건</span>
</nav>

<div class="container-fluid py-3 px-4">

  <!-- KPI 카드 -->
  <div class="row g-3 mb-3">
    <div class="col-md-3">
      <div class="card stat-box h-100">
        <div class="text-muted small">전체 블록</div>
        <div class="fs-2 fw-bold text-dark">{stats['total']:,}</div>
      </div>
    </div>
    <div class="col-md-3">
      <div class="card stat-box h-100">
        <div class="text-muted small">배정 완료</div>
        <div class="fs-2 fw-bold text-success">{stats['assigned']:,}</div>
        <div class="text-muted small">{stats['assigned']/stats['total']*100:.1f}%</div>
      </div>
    </div>
    <div class="col-md-3">
      <div class="card stat-box h-100">
        <div class="text-muted small">미배정</div>
        <div class="fs-2 fw-bold text-warning">{stats['unassigned']:,}</div>
        <div class="text-muted small">{stats['unassigned']/stats['total']*100:.1f}%</div>
      </div>
    </div>
    <div class="col-md-3">
      <div class="card stat-box h-100">
        <div class="text-muted small">배정 규칙 수</div>
        <div class="fs-2 fw-bold text-primary">{len(stats['steps'])}</div>
        <div class="text-muted small">R2 ~ R12</div>
      </div>
    </div>
  </div>

  <!-- 탭 네비게이션 -->
  <ul class="nav nav-tabs mb-3" id="mainTab">
    <li class="nav-item"><button class="nav-link active" data-bs-toggle="tab" data-bs-target="#tab-flow">🔄 배정 플로우챠트</button></li>
    <li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#tab-steps">📊 단계별 배정 결과</button></li>
    <li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#tab-util">🏭 작업장별 조업도</button></li>
    <li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#tab-blocks">📋 블록 목록</button></li>
  </ul>

  <div class="tab-content">

    <!-- ───────────── Tab 1: 플로우챠트 ───────────── -->
    <div class="tab-pane fade show active" id="tab-flow">
      <div class="card p-4">
        <h5 class="mb-3">📌 배정 규칙 전체 플로우챠트</h5>
        <div class="alert alert-info small mb-3">
          🟡 노란 박스: 리포팅 단계 (배정 없음) &nbsp;|&nbsp;
          🔵 파란 박스: 글로벌 계산 단계 &nbsp;|&nbsp;
          🟢 초록 박스: 목표 산출 및 저장
        </div>
        <div class="mermaid" style="min-height:600px">
{MERMAID_CHART}
        </div>
        <hr class="my-4">
        <h6>규칙별 요약</h6>
        <div class="row g-2">
          <div class="col-md-6"><div class="p-2 border rounded small"><b>R0</b> 이관물량 처리가능여부: T부하 + H부하×1.2 vs 전체능력 비교 → 초과 월 경고</div></div>
          <div class="col-md-6"><div class="p-2 border rounded small"><b>R1</b> 목표물량 산출: 전체조업도 = Σ부하 ÷ Σ능력 → 각 작업장 배정목표 = 조업도 × 능력</div></div>
          <div class="col-md-6"><div class="p-2 border rounded small"><b>R2</b> H/Y9 블록: blk이 Y9로 시작하는 H행 → AREA15 (목표 무관)</div></div>
          <div class="col-md-6"><div class="p-2 border rounded small"><b>R3</b> T-중/대: h_t=T이고 stg=중 또는 대 → m_area 값으로 직접 배정</div></div>
          <div class="col-md-6"><div class="p-2 border rounded small"><b>R4</b> T-소/동일: h_t=T, stg=소, area_prior_1=동일 → prt_area 배정</div></div>
          <div class="col-md-6"><div class="p-2 border rounded small"><b>R5</b> H-소 처리가능여부: H-소 부하 vs AREA3~5 잔여능력 비교 → 초과 월 경고</div></div>
          <div class="col-md-6"><div class="p-2 border rounded small"><b>R6</b> H-소: m_stdt 오름차순 정렬 후 AREA3→4→5 순환. 목표 도달 area 건너뜀</div></div>
          <div class="col-md-6"><div class="p-2 border rounded small"><b>R7</b> H-중-jig_F: Part1=특정블록 무조건 AREA2 / Part2=나머지 jig_F 목표까지</div></div>
          <div class="col-md-6"><div class="p-2 border rounded small"><b>R8</b> H-중-jig_D/L: H·E11·E51 시작 제외. AREA1→2→6 순환 배정</div></div>
          <div class="col-md-6"><div class="p-2 border rounded small"><b>R9</b> 전문화 Family: (pjt+blk앞3자리)로 그룹. P말미→AREA7/9, S말미→AREA8/10</div></div>
          <div class="col-md-6"><div class="p-2 border rounded small"><b>R10</b> H-대-jig_D/L: 씨드블록과 동일 (pjt+blk앞3자리+말미) 그룹 → AREA12</div></div>
          <div class="col-md-6"><div class="p-2 border rounded small"><b>R11</b> 잔여: PS페어→(7+8)(9+10)(13+14) / 단독→AREA7~14 first-fit / 초과시 개별</div></div>
          <div class="col-md-6"><div class="p-2 border rounded small"><b>R12</b> H-dir: blk 4번째 자리가 9 또는 H → prt_area로 강제 override 배정</div></div>
        </div>
      </div>
    </div>

    <!-- ───────────── Tab 2: 단계별 결과 ───────────── -->
    <div class="tab-pane fade" id="tab-steps">
      <div class="row g-3">
        <div class="col-lg-5">
          <div class="card p-3 h-100">
            <h6 class="mb-3">단계별 배정 건수</h6>
            <canvas id="stepChart" height="320"></canvas>
          </div>
        </div>
        <div class="col-lg-7">
          <div class="card p-3 h-100">
            <h6 class="mb-3">단계별 배정 상세</h6>
            <table class="table table-sm table-hover mb-0">
              <thead class="table-light">
                <tr>
                  <th>규칙</th><th>설명</th><th class="text-end">배정</th>
                  <th class="text-end">누적</th><th>비율</th><th>배정 작업장</th>
                </tr>
              </thead>
              <tbody>{step_rows}</tbody>
            </table>
          </div>
        </div>
      </div>

      <div class="card mt-3 p-3">
        <h6 class="mb-3">작업장별 배정 건수</h6>
        <canvas id="areaChart" height="120"></canvas>
      </div>
    </div>

    <!-- ───────────── Tab 3: 조업도 ───────────── -->
    <div class="tab-pane fade" id="tab-util">
      <div class="card p-3">
        <div class="d-flex justify-content-between align-items-center mb-2">
          <h6 class="mb-0">작업장별 월별 조업도 (배정부하 / 능력 × 100%)</h6>
          <div class="small">
            <span class="legend-item" style="background:#d1ecf1"></span>0~60%&nbsp;
            <span class="legend-item" style="background:#78c98a"></span>60~90%&nbsp;
            <span class="legend-item" style="background:#ffc107"></span>90~110%&nbsp;
            <span class="legend-item" style="background:#dc3545"></span>110%+
          </div>
        </div>
        <div class="table-responsive">
          <table class="table table-sm table-bordered mb-2" style="font-size:12px">
            <thead class="table-dark">
              <tr><th>작업장</th>{MONTHS_HDR}<th>연평균</th></tr>
            </thead>
            <tbody>{util_rows}</tbody>
          </table>
        </div>
        <div class="small text-muted">※ 셀에 마우스를 올리면 목표 조업도 확인 가능</div>
        <div class="mt-3">
          <canvas id="utilChart" height="120"></canvas>
        </div>
      </div>
    </div>

    <!-- ───────────── Tab 4: 블록 목록 ───────────── -->
    <div class="tab-pane fade" id="tab-blocks">
      <div class="filter-bar d-flex gap-3 flex-wrap align-items-center">
        <div>
          <label class="form-label mb-0 small fw-semibold">H/T</label>
          <select id="fHT" class="form-select form-select-sm" style="width:90px">
            <option value="">전체</option>
            <option>H</option><option>T</option>
          </select>
        </div>
        <div>
          <label class="form-label mb-0 small fw-semibold">규격</label>
          <select id="fStg" class="form-select form-select-sm" style="width:90px">
            <option value="">전체</option>
            <option>대</option><option>중</option><option>소</option>
          </select>
        </div>
        <div>
          <label class="form-label mb-0 small fw-semibold">배정규칙</label>
          <select id="fRule" class="form-select form-select-sm" style="width:100px">
            <option value="">전체</option>
            <option>R2</option><option>R3</option><option>R4</option>
            <option>R6</option><option>R7</option><option>R8</option>
            <option>R9</option><option>R10</option><option>R11</option><option>R12</option>
            <option value="NONE">미배정</option>
          </select>
        </div>
        <div>
          <label class="form-label mb-0 small fw-semibold">작업장</label>
          <select id="fArea" class="form-select form-select-sm" style="width:120px">
            <option value="">전체</option>
          </select>
        </div>
        <div class="ms-auto">
          <input id="fSearch" type="text" class="form-control form-control-sm" placeholder="🔍 blk 검색..." style="width:160px">
        </div>
        <div><span id="rowCount" class="badge bg-secondary">0건</span></div>
      </div>

      <div class="card p-0">
        <div style="max-height:520px;overflow-y:auto">
          <table class="table table-sm table-hover mb-0" id="blockTable">
            <thead class="table-dark sticky-top">
              <tr>
                <th>pjt</th><th>h_t</th><th>blk</th><th>stg</th>
                <th>m_stdt</th><th>m_fndt</th><th>m_area</th>
                <th>load_mh</th><th>assigned_area</th><th>rule</th>
              </tr>
            </thead>
            <tbody id="blockBody"></tbody>
          </table>
        </div>
      </div>
    </div>

  </div><!-- tab-content -->
</div><!-- container -->

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
<script>
// ── 임베디드 데이터 ──────────────────────────────────────────────────────────
const STATS  = {json.dumps(stats, ensure_ascii=False)};
const BLOCKS = {json.dumps(blocks_json, ensure_ascii=False)};

// ── Chart: 단계별 배정 건수 ─────────────────────────────────────────────────
const stepLabels = STATS.steps.map(s => s.rule);
const stepData   = STATS.steps.map(s => s.count);
new Chart(document.getElementById('stepChart'), {{
  type: 'bar',
  data: {{
    labels: stepLabels,
    datasets: [{{ label: '배정 건수', data: stepData,
      backgroundColor: stepData.map(v =>
        v > 100 ? '#0d6efd' : v > 50 ? '#198754' : v > 10 ? '#fd7e14' : '#6c757d'
      ), borderRadius: 4
    }}]
  }},
  options: {{ plugins:{{legend:{{display:false}}}}, scales:{{y:{{beginAtZero:true}}}} }}
}});

// ── Chart: 작업장별 배정 건수 ────────────────────────────────────────────────
const areaCounts = {{}};
BLOCKS.forEach(b => {{
  const a = b.assigned_area || '미배정';
  areaCounts[a] = (areaCounts[a] || 0) + 1;
}});
const areaLabels = Object.keys(areaCounts).sort();
const areaData   = areaLabels.map(a => areaCounts[a]);
new Chart(document.getElementById('areaChart'), {{
  type: 'bar',
  data: {{
    labels: areaLabels,
    datasets: [{{ label: '배정 건수', data: areaData,
      backgroundColor: '#0d6efd', borderRadius: 4
    }}]
  }},
  options: {{ plugins:{{legend:{{display:false}}}}, scales:{{y:{{beginAtZero:true}}}} }}
}});

// ── Chart: 조업도 월별 추이 (전체평균) ───────────────────────────────────────
const months12 = {json.dumps([f'2026-{m:02d}' for m in range(1,13)])};
const areas    = Object.keys(STATS.utilization).sort();
const utilDs   = areas.map((area, i) => ({{
  label: area,
  data: months12.map(m => STATS.utilization[area][m] || 0),
  borderWidth: 1.5, pointRadius: 2, fill: false,
  borderColor: `hsl(${{(i * 360/areas.length).toFixed(0)}},65%,50%)`
}}));
new Chart(document.getElementById('utilChart'), {{
  type: 'line',
  data: {{ labels: months12, datasets: utilDs }},
  options: {{
    plugins: {{ legend: {{ position:'right', labels:{{ boxWidth:10, font:{{size:10}} }} }} }},
    scales: {{ y: {{ beginAtZero:true, title:{{display:true, text:'조업도(%)' }} }} }}
  }}
}});

// ── 작업장 필터 셀렉트 채우기 ────────────────────────────────────────────────
const fArea = document.getElementById('fArea');
[...new Set(BLOCKS.map(b => b.assigned_area).filter(Boolean))].sort()
  .forEach(a => {{ const o = new Option(a,a); fArea.add(o); }});

// ── 블록 목록 렌더링 ─────────────────────────────────────────────────────────
let filteredBlocks = [...BLOCKS];

function renderTable() {{
  const tbody = document.getElementById('blockBody');
  const ht    = document.getElementById('fHT').value;
  const stg   = document.getElementById('fStg').value;
  const rule  = document.getElementById('fRule').value;
  const area  = document.getElementById('fArea').value;
  const srch  = document.getElementById('fSearch').value.toLowerCase();

  filteredBlocks = BLOCKS.filter(b => {{
    if (ht   && b.h_t !== ht) return false;
    if (stg  && b.stg !== stg) return false;
    if (rule === 'NONE' && b.rule_assigned !== '') return false;
    if (rule && rule !== 'NONE' && b.rule_assigned !== rule) return false;
    if (area && b.assigned_area !== area) return false;
    if (srch && !b.blk.toLowerCase().includes(srch)) return false;
    return true;
  }});

  document.getElementById('rowCount').textContent = filteredBlocks.length.toLocaleString() + '건';

  const display = filteredBlocks.slice(0, 1000);
  tbody.innerHTML = display.map(b => `
    <tr>
      <td>${{b.pjt}}</td>
      <td><span class="badge ${{b.h_t==='H'?'bg-primary':'bg-success'}}">${{b.h_t}}</span></td>
      <td class="font-monospace">${{b.blk}}</td>
      <td><span class="badge ${{b.stg==='대'?'bg-danger':b.stg==='중'?'bg-warning text-dark':'bg-secondary'}}">${{b.stg}}</span></td>
      <td>${{b.m_stdt}}</td>
      <td>${{b.m_fndt}}</td>
      <td>${{b.m_area}}</td>
      <td class="text-end">${{b.load_mh?Number(b.load_mh).toFixed(1):''}}</td>
      <td><span class="badge bg-info text-dark">${{b.assigned_area||'<span style="color:#aaa">미배정</span>'}}</span></td>
      <td><span class="badge bg-secondary">${{b.rule_assigned||'-'}}</span></td>
    </tr>`).join('');
  if (filteredBlocks.length > 1000) {{
    tbody.innerHTML += `<tr><td colspan="10" class="text-center text-muted">
      ... 상위 1,000건 표시 중 (전체 ${{filteredBlocks.length.toLocaleString()}}건) ...</td></tr>`;
  }}
}}

['fHT','fStg','fRule','fArea'].forEach(id =>
  document.getElementById(id).addEventListener('change', renderTable));
document.getElementById('fSearch').addEventListener('input', renderTable);
renderTable();
</script>
</body>
</html>"""

# ── 파일 저장 ──────────────────────────────────────────────────────────────────
Path(OUTPUT_PATH).write_text(HTML, encoding='utf-8')
print(f"✅ HTML 대시보드 생성 완료: {OUTPUT_PATH}")
print(f"   → 브라우저에서 열기: file:///{OUTPUT_PATH.replace(chr(92), '/')}")
