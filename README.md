# blk_assign_agent

> 제품과 작업장을 규칙 기반으로 자동 배정하고,  
> 월별·주별 조업도를 균등화하는 Python 기반 배정 엔진.  
> 실행 한 번으로 배정 결과(xlsx)와 웹 대시보드(html)를 동시에 생성한다.

---

## 목차
1. [시스템 개요](#1-시스템-개요)
2. [디렉터리 구조](#2-디렉터리-구조)
3. [아키텍처](#3-아키텍처)
4. [입력 데이터 명세](#4-입력-데이터-명세)
5. [배정 규칙 요약 (R0~R13)](#5-배정-규칙-요약-r0r13)
6. [설치 및 환경 설정](#6-설치-및-환경-설정)
7. [실행 방법](#7-실행-방법)
8. [출력 파일 설명](#8-출력-파일-설명)
9. [웹 대시보드 사용법](#9-웹-대시보드-사용법)
10. [주요 로직 설명](#10-주요-로직-설명)
11. [유지보수 가이드](#11-유지보수-가이드)
12. [향후 개선 과제](#12-향후-개선-과제)

---

## 1. 시스템 개요

| 항목 | 내용 |
|------|------|
| 목적 | 블록별 load_mh를 작업장 능력(area_capa)에 맞게 균등 배분 |
| 핵심 개념 | 목표조업도 = 전체부하 ÷ 전체능력, 작업장별 목표물량 = 목표조업도 × 작업장능력 |
| 부하 산출 | m_stdt ~ m_fndt 기간의 영업일에 load_mh를 일할 분배 → 월별·주별 부하 집계 |
| 배정 방식 | R0~R13 규칙을 순차 실행, 각 규칙은 독립적으로 수정 가능 |
| 출력 | blk_assign_result.xlsx + blk_assign_stats.json + blk_assign_report.html (실행 시 자동 생성) |

---

## 2. 디렉터리 구조

```
blk_assign/
│
├── 📄 README.md                  # 이 파일
│
├── 🔧 [배정 엔진 + 리포트]
│   ├── blk_assign_agent.py       # 핵심: 배정 로직 (R0~R13) + HTML 리포트 자동 생성
│   ├── blk_assign.py             # 구버전 배정 스크립트 (참고용)
│   └── app.py                    # Streamlit 대시보드 (배정 실행 + 인터랙티브 시각화)
│
├── 🗄️ [데이터 생성]
│   ├── blk_master_gen.py         # 테스트용 blk_master 데이터 생성기
│   └── blk_master_gen_spec.md    # blk_master_gen.py 컬럼별 생성 규칙 명세
│
├── 📊 [입력 데이터]
│   ├── blk_master.xlsx           # 블록 마스터 (배정 대상)
│   └── area_capa.xlsx            # 작업장별 월별 능력(MH)
│
├── 📁 [출력 파일]  ─── blk_assign_agent.py / app.py 실행 시 자동 생성
│   ├── blk_assign_result.xlsx    # 배정 결과 (assigned_area 포함)
│   ├── blk_assign_stats.json     # 배정 통계
│   └── blk_assign_report.html    # 정적 웹 대시보드 (브라우저에서 직접 열기)
│
└── 📋 [명세 문서]
    ├── blk_assign_spec_v1.0.md   # 배정 규칙 상세 명세 (Markdown)
    └── blk_assign_spec_v1.0.txt  # 배정 규칙 상세 명세 (Text)
```

---

## 3. 아키텍처

```
┌─────────────────────────────────────────────────────────────┐
│                        입력 데이터                           │
│  blk_master.xlsx          area_capa.xlsx                    │
│  (블록 마스터 - 배정 대상)  (작업장별 월별 능력)              │
└────────────────┬─────────────────────┬──────────────────────┘
                 │                     │
                 ▼                     ▼
┌─────────────────────────────────────────────────────────────┐
│                   blk_assign_agent.py                       │
│                                                             │
│  [글로벌] load_mh 일할 분배 → 월별/주별 부하 산출            │
│  [R1]  목표조업도 및 작업장별 목표물량 산출                   │
│  [R2~R13] 규칙 기반 순차 배정                                │
│                                                             │
│  ── 배정 완료 후 자동 실행 ──────────────────────────────    │
│  generate_html_report()                                     │
│  - Mermaid 플로우차트                                        │
│  - 단계별 배정 결과 차트                                     │
│  - 작업장별 조업도 히트맵                                    │
│  - 블록 목록 검색/필터                                       │
│  - Agent 실행 시점 + 페이지 로드 시점 타임스탬프             │
└──────┬──────────────┬──────────────────┬───────────────────┘
       │              │                  │
       ▼              ▼                  ▼
┌────────────┐ ┌─────────────┐ ┌──────────────────────┐
│result.xlsx │ │stats.json   │ │blk_assign_report.html │
│(배정 결과) │ │(배정 통계)  │ │(브라우저에서 직접 열기)│
└────────────┘ └─────────────┘ └──────────────────────┘
```

### 모듈 역할

| 파일 | 역할 | 실행 순서 |
|------|------|-----------|
| `blk_master_gen.py` | 테스트용 blk_master 데이터 생성 | ① (최초 1회) |
| `blk_assign_agent.py` | 배정 엔진 (R0~R13) + HTML 리포트 자동 생성 | ② |
| `app.py` | Streamlit 대시보드 — 배정 실행 + 인터랙티브 시각화 | ② (대안) |

> `blk_assign_agent.py`와 `app.py`는 같은 배정 로직을 공유한다.  
> CLI 실행 후 정적 HTML이 필요하면 `blk_assign_agent.py`,  
> 브라우저에서 실시간 필터/차트가 필요하면 `app.py`를 사용한다.

---

## 4. 입력 데이터 명세

### blk_master.xlsx — 주요 컬럼

| 컬럼 | 설명 | 예시 |
|------|------|------|
| `pjt` | 프로젝트 번호 | 1000~1020 |
| `ship_kind` | 선종 구분 | A, B, C, D |
| `h_t` | H(이관물량) / T(자체물량) | H, T |
| `blk` | 블록 코드 (5자리) | A123P |
| `stg` | 규격 (대/중/소) | 대, 중, 소 |
| `type` | 작업 유형 (T행만) | AA~DD (16종) |
| `jig` | 지그 유형 (H행만) | D, F, L, W |
| `load_mh` | 블록 부하 (Man-Hour) | 340.9 |
| `m_stdt` | 작업 착수일 | 2026-03-10 |
| `m_fndt` | 작업 완료일 | 2026-03-25 |
| `m_area` | 계획 작업장 | AREA5 |
| `prt_area` | 부모 블록 작업장 (T-소 행) | AREA3 |
| `area_prior_1~5` | 작업장 배정 우선순위 후보 | AREA1, 동일, ... |
| `assign_prior` | 배정 우선순위 번호 (1~16) | 3 |

### area_capa.xlsx — 구조

| 컬럼 | 설명 |
|------|------|
| 첫 번째 컬럼 | 작업장명 (AREA1~AREA14) |
| 1월~12월 컬럼 | 해당 월 작업 능력 (Man-Hour) |
| `assign_prior` 컬럼 | R13 배정 순서 (area_seq) |

---

## 5. 배정 규칙 요약 (R0~R13)

> 상세 명세는 [`blk_assign_spec_v1.0.md`](blk_assign_spec_v1.0.md) 참조

| 규칙 | 목적 | 대상 조건 | 배정 작업장 |
|------|------|-----------|------------|
| R0 | 이관물량 처리가능여부 리포팅 | 전체 | — (리포팅만) |
| R1 | 작업장별 월별 배정 목표물량 산출 | 전체 | — (기준 산출) |
| R2 | H/Y9 블록 | h_t=H, blk 앞2자리=Y9 | AREA15 |
| R3 | T-중/대 | h_t=T, stg∈{중,대} | m_area 값 |
| R4 | T-소/동일 | h_t=T, stg=소, area_prior_1=동일 | prt_area 값 |
| R5 | H-소 처리가능여부 리포팅 | h_t=H, stg=소 | — (리포팅만) |
| R6 | H-소 순환 배정 | h_t=H, stg=소 | AREA3→4→5 순환 |
| R7 | H-중 jig_F | h_t=H, stg=중, jig=F | AREA2 |
| R8 | H-중 jig_D/L/W | h_t=H, stg=중, jig∈{D,L,W} | AREA1→2→6 순환 |
| R9 | H 전문화 Family | h_t=H, stg∈{대,중}, blk패턴 | AREA7/8/9/10 |
| R10 | H-대 jig_D/L 그룹 | h_t=H, stg=대, jig∈{D,L} | AREA12 |
| R11 | H-대/중 잔여 | h_t=H, stg∈{대,중}, 미배정 | AREA7~14 |
| R12 | H-dir override | h_t=H, blk[3]∈{9,H} | prt_area 값 (override) |
| R13 | T-소 분산 배정 | h_t=T, stg=소, 미배정 | area_seq 순서대로 |

---

## 6. 설치 및 환경 설정

### 필수 패키지

```bash
# blk_assign_agent.py 실행 시
pip install pandas openpyxl numpy

# app.py (Streamlit 대시보드) 실행 시 추가 설치
pip install streamlit plotly
```

### Python 버전
- Python 3.9 이상 권장

### 파일 경로 설정
`blk_assign_agent.py` 및 `blk_assign.py` 상단의 `BASE_DIR`을 실제 파일 위치로 수정한다.  
`app.py`는 스크립트 위치를 자동 감지하므로 별도 수정이 필요 없다.

```python
# blk_assign_agent.py / blk_assign.py 상단 (Windows)
BASE_DIR        = r'D:\AI\blk_assign'   # ← 실제 경로로 변경
BLK_MASTER_PATH = f'{BASE_DIR}/blk_master.xlsx'
AREA_CAPA_PATH  = f'{BASE_DIR}/area_capa.xlsx'
RESULT_PATH     = f'{BASE_DIR}/blk_assign_result.xlsx'
STATS_PATH      = f'{BASE_DIR}/blk_assign_stats.json'
REPORT_PATH     = f'{BASE_DIR}/blk_assign_report.html'
```

> WSL2 환경에서 Windows 경로 접근 시: `BASE_DIR = '/mnt/d/AI/blk_assign'`

---

## 7. 실행 방법

### Step 1 — (선택) 테스트 데이터 생성

실제 blk_master가 없을 경우 테스트 데이터를 생성한다.

```bash
python blk_master_gen.py
```

- 생성 결과: `blk_master.xlsx` (3,200행 × 24열)

### Step 2 — 배정 엔진 실행

```bash
python blk_assign_agent.py
```

- 실행 시간: 약 10~30초 (데이터 규모에 따라 상이)
- 출력 파일: `blk_assign_result.xlsx`, `blk_assign_stats.json`, `blk_assign_report.html`
- 콘솔 출력: R0~R13 단계별 배정 결과, 조업도 요약, HTML 생성 완료 메시지

**콘솔 출력 예시:**
```
[R2] H/Y9 → AREA15: 1건
[R3] T-중/대 → m_area: 122건
...
[R13] T-소 분산배정 합계: 462건
전체: 3,200건 | 배정완료: 2,654건 | 미배정: 546건
📊 통계 저장 완료: blk_assign_stats.json
✅ 저장 완료: blk_assign_result.xlsx
🌐 HTML 리포트 생성 완료: blk_assign_report.html
```

### Step 3 — 웹 대시보드 확인

Step 2 완료 후 자동 생성된 `blk_assign_report.html`을 브라우저에서 직접 열면 된다.

```bash
# Windows
start blk_assign_report.html

# macOS
open blk_assign_report.html
```

> **별도 실행 불필요** — `blk_assign_web.py`는 `blk_assign_agent.py`에 통합되었다.

---

## 8. 출력 파일 설명

### blk_assign_result.xlsx
blk_master에 다음 2개 컬럼이 추가된 결과 파일

| 추가 컬럼 | 설명 |
|-----------|------|
| `assigned_area` | 배정된 작업장명 (미배정 시 공란) |
| `rule_assigned` | 배정에 사용된 규칙 번호 (R2~R13) |

### blk_assign_stats.json
웹 대시보드 렌더링에 사용되는 통계 데이터

```json
{
  "total": 3200,
  "assigned": 2654,
  "unassigned": 546,
  "steps": [
    {"rule": "R2", "label": "H/Y9 → AREA15", "count": 1, "areas": {...}},
    ...
  ],
  "utilization": {"AREA1": {"2026-01": 63.8, ...}, ...},
  "target_util":  {"AREA1": {"2026-01": 68.1, ...}, ...}
}
```

---

## 9. 웹 대시보드 사용법

`blk_assign_report.html`을 브라우저에서 열면 4개 탭으로 구성된 대시보드가 표시된다.

네비게이션 바 우측에 두 가지 타임스탬프가 표시된다.

| 타임스탬프 | 의미 | 갱신 시점 |
|-----------|------|-----------|
| ⚙️ Agent 실행 | `blk_assign_agent.py` 실행 시각 | HTML 생성 시 고정 |
| 🌐 페이지 로드 | 브라우저가 HTML을 로드한 시각 | 탭 열 때마다 갱신 |

| 탭 | 내용 |
|----|------|
| 🔄 배정 플로우차트 | R0~R13 전체 배정 흐름 다이어그램 + 규칙별 설명 카드 |
| 📊 단계별 배정 결과 | 규칙별 배정 건수 막대차트, 누적 배정 현황, 작업장별 분포 |
| 🏭 작업장별 조업도 | 월별 조업도 색상 히트맵 (목표 대비 실적 비교), 월별 추이 라인차트 |
| 📋 블록 목록 | h_t/stg/규칙/작업장/blk 코드 기준 필터·검색, 상위 1,000건 표시 |

> 외부 CDN(Bootstrap, Chart.js, Mermaid.js)을 사용하므로 **인터넷 연결이 필요**하다.

---

## 10. 주요 로직 설명

### 부하 일할 분배
```python
daily_mh = load_mh / len(workdays)   # 작업기간 내 영업일 수로 나눔
for d in workdays:
    monthly[(d.year, d.month)] += daily_mh
    weekly[(iso.year, iso.week)]  += daily_mh
```

### 목표물량 도달 판정 (skip_streak 방식)
단순 break 대신, 연속 50건 거부 시에만 중단하여 월별 부하 분포가 다른 블록도 배정 기회를 갖도록 한다.
```python
skip_streak = 0
for idx in candidates:
    if can_assign(area, load_dists[idx], cl, tgt):
        assign_block(...)
        skip_streak = 0
    else:
        skip_streak += 1
        if skip_streak >= 50:
            break
```

### can_assign — 월별 목표 초과 여부 확인
```python
def can_assign(area, dist, cl, tgt) -> bool:
    for (yr, mo), blk_load in dist['monthly'].items():
        if cl[area][(yr, mo)] + blk_load > tgt[area][(yr, mo)]:
            return False
    return True
```

---

## 11. 유지보수 가이드

### 규칙 수정 방법

각 규칙은 독립적인 함수로 분리되어 있으므로 해당 함수만 수정하면 된다.

```
r0_report_transfer_capacity()   → R0 리포팅 기준 변경
r1_calc_target_loads()          → 목표조업도 산출 방식 변경
r2_assign_y9()                  → Y9 조건 또는 배정 작업장 변경
r3_assign_t_mid_large()         → T-중/대 배정 로직 변경
r4_assign_t_small_dongil()      → T-소/동일 배정 로직 변경
r6_assign_h_small()             → H-소 순환 작업장 변경
r7_assign_h_mid_jig_f()         → jig_F 특수 블록 조건 변경
r8_assign_h_mid_jig_dl()        → jig 종류 추가/변경
r9_assign_specialized_family()  → Family 그룹 조건 변경
r10_assign_h_large_jig_dl()     → 씨드 조건 또는 배정 작업장 변경
r11_assign_h_remaining()        → 잔여 물량 배정 순서 변경
r12_assign_h_dir()              → override 조건 변경
r13_assign_t_small_dist()       → T-소 분산 배정 우선순위 변경
```

### 새 규칙 추가 방법

1. `blk_assign_agent.py`에 `r14_xxx()` 함수 작성
2. 함수 내 첫 줄에 `set_rule('R14')` 추가
3. `RULE_LABELS` 딕셔너리에 `'R14': '설명'` 추가
4. `if __name__ == '__main__':` 블록에서 `r13_...` 호출 다음에 `r14_xxx(...)` 추가

### 공휴일 업데이트
`blk_assign_agent.py` 상단의 `KR_HOLIDAYS` 집합에 연도별 공휴일을 추가한다.

```python
KR_HOLIDAYS = {
    date(2026, 1, 1),
    # ... 2026년 공휴일
    date(2027, 1, 1),   # ← 2027년 추가 시
    # ...
}
```

### 작업장 추가/변경
- 작업장 추가: `area_capa.xlsx`에 행 추가 후 `assign_prior` 값 설정
- 배정 순서 변경: `area_capa.xlsx`의 `assign_prior` 컬럼 값 수정
- AREA15처럼 capa 없는 특수 작업장: `AREA_Y9 = 'AREA15'` 상수 확인

---

## 12. 향후 개선 과제

| 우선순위 | 항목 | 설명 |
|----------|------|------|
| 🔴 높음 | T-소 미배정 처리 | area_prior_1≠동일인 T-소 잔여 배정 규칙 추가 필요 |
| 🔴 높음 | H-대 jig_F/W 규칙 | H-대이고 jig=F 또는 W인 블록 배정 규칙 미정의 |
| 🟡 보통 | 주별 평활화 강화 | 인접 주 여유도를 작업장 선택 점수에 반영 |
| 🟡 보통 | 웹 대시보드 서버화 | Flask 기반 실시간 필터링 서버로 전환 (현재는 정적 HTML) |
| 🟢 낮음 | 2027년 공휴일 추가 | m_fndt가 2027년에 걸치는 블록 처리 |
| 🟢 낮음 | 배정 결과 Excel 포맷팅 | 작업장별 색상 구분, 조건부 서식 적용 |

---

## 기여 방법

1. 이 저장소를 Fork
2. 기능 브랜치 생성 (`git checkout -b feature/새기능`)
3. 변경 사항 커밋 (`git commit -m 'Add: 새기능 설명'`)
4. 브랜치에 Push (`git push origin feature/새기능`)
5. Pull Request 생성

---

## 라이선스

본 프로젝트는 내부 업무용 시스템으로 별도 라이선스를 명시하지 않습니다.
