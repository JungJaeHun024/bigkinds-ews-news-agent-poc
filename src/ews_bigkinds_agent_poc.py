"""
BigKinds 기반 EWS 언론보도 AI Agent PoC

This script is exported from the GitHub-ready notebook.
Recommended execution environment: Google Colab / Jupyter Notebook.
For portfolio review, see `ews_bigkinds_agent_poc_github.ipynb`.
"""


# %% [markdown]
"""
# BigKinds 기반 EWS 언론보도 AI Agent PoC

이 노트북은 BigKinds 뉴스 Excel 데이터를 기반으로 건설산업 조기경보시스템(EWS)의 비정형 언론보도 데이터를 구조화할 수 있는지 검증하는 PoC입니다.

## PoC Scope
- **목적**: 완성형 뉴스 검색엔진 개발이 아니라, BigKinds Excel 기반 언론보도 데이터를 EWS 도메인별 위험 후보 데이터셋으로 구조화할 수 있는지 검증
- **도메인**: 건설경기, 금융시장, 건설자재, 노동시장
- **핵심 흐름**: Excel Ingestion → DB 적재 → 도메인 분류 → 위험 이벤트 후보 생성 → LLM 최종 판단 → Excel Report
- **LLM 역할**: 전체 기사 전수 판단이 아니라, Rule 기반 후보에 대한 최종 판단 보조
- **수집 방식**: BigKinds 웹 검색 결과 Excel 다운로드 기반 반자동 수집

## Portfolio Points
1. 산업 도메인 특화 키워드/이벤트 사전 설계
2. LLM timeout 대응을 위한 batch 축소, 빠른 모델 옵션, streaming 호출
3. 수집 가능성 판단 목적에 맞춘 PoC 범위 통제

## Before Running
1. Colab에 BigKinds Excel 파일을 업로드합니다.
2. `BIGKINDS_FILE_PATH` 환경변수를 지정하거나 Cell 2의 `FILE_PATH`를 수정합니다.
3. NVIDIA API Key는 Cell 2에서 `getpass()`로 입력합니다.

"""

# %% [markdown]
"""
## 1. Environment Setup

필요 패키지를 설치합니다. Colab 기준으로 실행하도록 구성했습니다.

"""

# %%
# =========================================
# Cell 1. 패키지 설치
# =========================================
# Install before running if needed:
# pip install -q pandas openpyxl tqdm requests langgraph chromadb sentence-transformers

# %% [markdown]
"""
## 2. Configuration

파일 경로, NVIDIA API 설정, LLM batch/streaming 정책을 관리합니다. API Key는 `getpass()`로 입력하여 저장소에 남기지 않습니다.

"""

# %%
# =========================================
# Cell 2. 기본 설정 및 API Key 입력
# =========================================

import os
import re
import json
import time
import uuid
import sqlite3
import traceback
import warnings
import requests

from datetime import datetime
from typing import Any, Dict, List, Optional, TypedDict

import pandas as pd
from tqdm.auto import tqdm
from getpass import getpass

warnings.filterwarnings("ignore")

# =========================================
# 1. 경로 설정
# =========================================
# Colab에 업로드한 빅카인즈 파일명에 맞게 FILE_PATH만 필요 시 수정
FILE_PATH = os.getenv("BIGKINDS_FILE_PATH", "/content/NewsResult_20260517-20260518.xlsx")

DB_PATH = "/content/ews_bigkinds_agent.db"
OUTPUT_EXCEL_PATH = "/content/ews_bigkinds_agent_report.xlsx"
VECTOR_DB_PATH = "/content/ews_bigkinds_chroma"

# =========================================
# 2. NVIDIA API Key 설정
# =========================================
if not os.environ.get("NVIDIA_API_KEY"):
    os.environ["NVIDIA_API_KEY"] = getpass("NVIDIA_API_KEY 입력: ")

NVIDIA_API_KEY = os.environ.get("NVIDIA_API_KEY", "")
NVIDIA_API_URL = "https://integrate.api.nvidia.com/v1/chat/completions"

# =========================================
# 3. NVIDIA 모델 선택 옵션
# =========================================
# 강한 모델: 품질은 좋지만 timeout 가능성 큼
NVIDIA_MODEL_STRONG = "meta/llama-3.1-70b-instruct"

# 빠른 모델: Colab / PoC / 무료 API 환경에서 안정성 우선
NVIDIA_MODEL_FAST = "meta/llama-3.1-8b-instruct"

# 기본 사용 모델
# Cell 11, Cell 13 안정성을 위해 기본값은 FAST로 둔다.
NVIDIA_MODEL = NVIDIA_MODEL_FAST

# =========================================
# 4. NVIDIA API 호출 제한 및 안정화 설정
# =========================================
# NVIDIA 개인 개발 환경 기준 분당 40회 요청 제한을 가정
NVIDIA_RPM_LIMIT = 40

# 실제 호출은 여유 있게 분당 24회로 제한
SAFE_RPM_LIMIT = 24

# 분당 24회 기준 약 2.5초 간격
NVIDIA_CALL_INTERVAL_SEC = 60 / SAFE_RPM_LIMIT

# Streaming을 쓰더라도 연결 대기 시간이 필요하므로 90초로 설정
NVIDIA_TIMEOUT_SEC = 90

# 재시도는 1회만 수행
# 너무 많이 재시도하면 Cell 11 / Cell 13 실행 시간이 길어짐
NVIDIA_MAX_RETRIES = 1

# Streaming 응답 사용 여부
# Cell 4에서 streaming 지원 함수로 교체하면 이 값이 사용됨
USE_NVIDIA_STREAMING = True

# =========================================
# 5. LLM 사용 정책
# =========================================
# Query Planner는 우선 Rule 기반으로 빠르게 처리
# 자연어 질의 해석까지 LLM으로 하고 싶으면 True로 변경 가능
USE_LLM_QUERY_PLANNER = False

# Query Planner 실패 시 Rule Planner fallback 허용
ALLOW_RULE_PLANNER_FALLBACK = True

# 최종 위험 판단은 LLM이 수행
USE_LLM_FINAL_JUDGEMENT = True

# 최종 판단 실패 시 Rule 결과를 최종 판단으로 둔갑시키지 않음
ALLOW_RULE_FINAL_FALLBACK = False

# =========================================
# 6. Final Risk Judgement 범위 설정
# =========================================
# 기본 모니터링 / 질의 응답 모두에서 LLM에 넘길 후보 수 제한
# 초기 안정성 테스트는 6건 권장
LLM_FINAL_TOP_N = 6

# 한 번의 LLM 호출에 넣을 기사 수
# timeout 방지를 위해 2건 단위로 분할
LLM_BATCH_SIZE = 2

# 기사별 본문/근거 텍스트 최대 길이
# 길수록 판단 품질은 좋아질 수 있지만 timeout 가능성이 커짐
LLM_CONTENT_MAX_CHARS = 120

# =========================================
# 7. 기본 실행 옵션
# =========================================
# Cell 11 기본 모니터링에서는 더 작은 범위로 LLM 판단
DEFAULT_MONITORING_LLM_TOP_N = 4
DEFAULT_MONITORING_LLM_BATCH_SIZE = 2

# Cell 13 직접 자연어 질의에서는 기본보다 조금 더 넓게 판단
INTERACTIVE_QUERY_LLM_TOP_N = 6
INTERACTIVE_QUERY_LLM_BATCH_SIZE = 2

# =========================================
# 8. 출력 확인
# =========================================
print("=" * 70)
print("Cell 2 설정 완료 — EWS BigKinds LangGraph Agent")
print("-" * 70)
print(f"FILE_PATH                  : {FILE_PATH}")
print(f"DB_PATH                    : {DB_PATH}")
print(f"OUTPUT_EXCEL_PATH          : {OUTPUT_EXCEL_PATH}")
print(f"VECTOR_DB_PATH             : {VECTOR_DB_PATH}")
print("-" * 70)
print(f"NVIDIA_API_URL             : {NVIDIA_API_URL}")
print(f"NVIDIA_MODEL               : {NVIDIA_MODEL}")
print(f"NVIDIA_MODEL_FAST          : {NVIDIA_MODEL_FAST}")
print(f"NVIDIA_MODEL_STRONG        : {NVIDIA_MODEL_STRONG}")
print(f"USE_NVIDIA_STREAMING       : {USE_NVIDIA_STREAMING}")
print("-" * 70)
print(f"NVIDIA_RPM_LIMIT           : {NVIDIA_RPM_LIMIT}")
print(f"SAFE_RPM_LIMIT             : {SAFE_RPM_LIMIT}")
print(f"NVIDIA_CALL_INTERVAL_SEC   : {NVIDIA_CALL_INTERVAL_SEC:.2f}")
print(f"NVIDIA_TIMEOUT_SEC         : {NVIDIA_TIMEOUT_SEC}")
print(f"NVIDIA_MAX_RETRIES         : {NVIDIA_MAX_RETRIES}")
print("-" * 70)
print(f"USE_LLM_QUERY_PLANNER      : {USE_LLM_QUERY_PLANNER}")
print(f"USE_LLM_FINAL_JUDGEMENT    : {USE_LLM_FINAL_JUDGEMENT}")
print(f"ALLOW_RULE_FINAL_FALLBACK  : {ALLOW_RULE_FINAL_FALLBACK}")
print("-" * 70)
print(f"LLM_FINAL_TOP_N            : {LLM_FINAL_TOP_N}")
print(f"LLM_BATCH_SIZE             : {LLM_BATCH_SIZE}")
print(f"LLM_CONTENT_MAX_CHARS      : {LLM_CONTENT_MAX_CHARS}")
print(f"DEFAULT_MONITORING_TOP_N   : {DEFAULT_MONITORING_LLM_TOP_N}")
print(f"INTERACTIVE_QUERY_TOP_N    : {INTERACTIVE_QUERY_LLM_TOP_N}")
print("=" * 70)

# %% [markdown]
"""
## 3. Domain & Risk Event Dictionary

건설경기·금융시장·건설자재·노동시장 도메인 사전과 위험 이벤트 사전을 정의합니다.

"""

# %%
# =========================================
# Cell 3. 도메인 및 위험 이벤트 사전
# =========================================
DOMAIN_KEYWORDS = {
    "건설경기": [
        "건설경기","건설투자","건설수주","수주","착공","분양","미분양","악성 미분양",
        "주택시장","부동산 경기","건설사","시행사","시공사","공사비","SOC","주택공급",
        "아파트","부동산","재건축","재개발","건설업","건설산업"
    ],
    "금융시장": [
        "금융시장","금리","기준금리","국채금리","환율","원달러","달러","회사채","CP","단기사채",
        "자금시장","유동성","가계부채","주담대","대출","연체율","저축은행","상호금융",
        "부동산 PF","PF","금융위","금감원","한국은행","FOMC","증시","코스피","코스닥"
    ],
    "건설자재": [
        "철근","시멘트","레미콘","골재","자재비","원자재","건설자재","철강","유연탄",
        "공급망","수급","납품","가격 상승","원가","공사원가","자재"
    ],
    "노동시장": [
        "고용","임금","노동","파업","인력난","건설노동","외국인력","외국인 노동자",
        "노조","중대재해","산재","근로자","일자리","채용","실업","노동시장"
    ]
}

RISK_EVENTS = {
    "부동산 PF 리스크":        ["부동산 PF","PF","프로젝트파이낸싱","브릿지론","사업장 정리","PF 부실","PF 대출"],
    "건설수주 감소":           ["건설수주 감소","수주 감소","수주 부진","발주 감소","수주 위축","건설수주"],
    "미분양 증가":             ["미분양","악성 미분양","준공 후 미분양","분양 부진"],
    "착공 지연":               ["착공 지연","착공 감소","인허가 지연","공사 지연","공기 지연"],
    "공사비·자재비 상승":      ["공사비","자재비","인건비","공사원가","원자재","가격 상승","원가 상승"],
    "금리·통화정책":           ["금리","기준금리","국채금리","FOMC","통화정책","금리 인상","금리 인하"],
    "환율·외환 변동":          ["환율","원달러","달러 강세","외환시장","외환","원화 약세"],
    "회사채·CP·자금시장":      ["회사채","CP","단기사채","신용스프레드","자금시장","자금조달","유동성"],
    "가계부채·대출":           ["가계부채","주담대","대출 규제","DSR","대출","부채"],
    "금융기관 건전성":         ["연체율","고정이하여신","충당금","저축은행","상호금융","은행 건전성","건전성"],
    "자재 공급망 불안":        ["철근","시멘트","레미콘","건설자재","공급망","수급 차질","납품 차질"],
    "노동·임금·인력 리스크":  ["인력난","임금","파업","노조","외국인력","건설노동","중대재해","노동시장"],
    "정책·시장안정":           ["시장안정","유동성 공급","정책금융","금융위","금감원","대책","관계기관"]
}

DOMAIN_EVENT_HINTS = {
    "건설경기": ["부동산 PF 리스크","건설수주 감소","미분양 증가","착공 지연","공사비·자재비 상승"],
    "금융시장": ["부동산 PF 리스크","금리·통화정책","환율·외환 변동","회사채·CP·자금시장","가계부채·대출","금융기관 건전성","정책·시장안정"],
    "건설자재": ["공사비·자재비 상승","자재 공급망 불안"],
    "노동시장": ["노동·임금·인력 리스크"]
}

print(f"도메인 사전: {len(DOMAIN_KEYWORDS)}개 도메인 / 위험 이벤트: {len(RISK_EVENTS)}개")

# %% [markdown]
"""
## 4. Utility Functions & NVIDIA Safe Caller

공통 유틸리티와 NVIDIA LLM 안전 호출 래퍼입니다. timeout 대응을 위해 streaming, rate limit, retry를 반영했습니다.

"""

# %%
# =========================================
# Cell 4. 공통 유틸리티 및 NVIDIA 안전 호출 래퍼
# =========================================
LAST_NVIDIA_CALL_TS = 0


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def safe_str(x: Any) -> str:
    if x is None:
        return ""
    if isinstance(x, float) and pd.isna(x):
        return ""
    return str(x)


def normalize_text(text: Any) -> str:
    text = safe_str(text)
    text = text.replace("\r"," ").replace("\n"," ").replace("\t"," ")
    return re.sub(r"\s+", " ", text).strip()


def make_article_text(article: Dict[str, Any]) -> str:
    keys = ["title","keywords","features","content","category_1","category_2","category_3"]
    return normalize_text(" ".join(safe_str(article.get(k,"")) for k in keys))


def safe_json_loads(text: str, default: Any = None) -> Any:
    if default is None:
        default = {}
    if not text:
        return default
    raw = text.strip().replace("```json","").replace("```","").strip()
    try:
        return json.loads(raw)
    except Exception:
        pass
    match = re.search(r"(\{.*\}|\[.*\])", raw, flags=re.DOTALL)
    if match:
        try:
            return json.loads(match.group(1))
        except Exception:
            return default
    return default


def extract_evidence_sentence_tool(text: str, matched_keywords: List[str], max_len: int = 260) -> str:
    text = normalize_text(text)
    if not text:
        return ""
    for kw in [k for k in matched_keywords if k]:
        idx = text.find(kw)
        if idx >= 0:
            return text[max(0, idx-90):min(len(text), idx+max_len)].strip()
    return text[:max_len].strip()


def wait_for_nvidia_rate_limit():
    global LAST_NVIDIA_CALL_TS
    wait_sec = max(0, NVIDIA_CALL_INTERVAL_SEC - (time.time() - LAST_NVIDIA_CALL_TS))
    if wait_sec > 0:
        time.sleep(wait_sec)
    LAST_NVIDIA_CALL_TS = time.time()


def call_nvidia_llm_safe(
    messages: List[Dict[str, str]],
    max_tokens: int = 900,
    temperature: float = 0.1,
    timeout_sec: Optional[int] = None,
    model: Optional[str] = None,
    stream: Optional[bool] = None
) -> Dict[str, Any]:
    """
    NVIDIA LLM 안전 호출 함수.

    개선 사항:
    1. streaming 응답 지원
    2. 빠른 모델 선택 가능
    3. rate limit 적용
    4. timeout / retry 적용
    5. 실패 시 전체 Agent를 중단하지 않고 상태 반환
    """

    _timeout = timeout_sec if timeout_sec is not None else NVIDIA_TIMEOUT_SEC
    _model = model or NVIDIA_MODEL
    _stream = USE_NVIDIA_STREAMING if stream is None else stream

    if not NVIDIA_API_KEY:
        return {
            "success": False,
            "content": None,
            "error": "NVIDIA_API_KEY 없음",
            "attempt": 0,
            "model": _model
        }

    headers = {
        "Authorization": f"Bearer {NVIDIA_API_KEY}",
        "Content-Type": "application/json"
    }

    payload = {
        "model": _model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
        "stream": _stream
    }

    last_error = None

    for attempt in range(NVIDIA_MAX_RETRIES + 1):
        try:
            wait_for_nvidia_rate_limit()

            if _stream:
                resp = requests.post(
                    NVIDIA_API_URL,
                    headers=headers,
                    json=payload,
                    timeout=_timeout,
                    stream=True
                )
                resp.raise_for_status()

                chunks = []
                for line in resp.iter_lines(decode_unicode=True):
                    if not line:
                        continue

                    if line.startswith("data: "):
                        data = line[len("data: "):].strip()

                        if data == "[DONE]":
                            break

                        try:
                            obj = json.loads(data)
                            delta = obj.get("choices", [{}])[0].get("delta", {})
                            content_piece = delta.get("content", "")
                            if content_piece:
                                chunks.append(content_piece)
                        except Exception:
                            continue

                content = "".join(chunks).strip()

                if not content:
                    raise ValueError("Streaming 응답은 성공했지만 content가 비어 있음")

                return {
                    "success": True,
                    "content": content,
                    "error": None,
                    "attempt": attempt + 1,
                    "model": _model,
                    "stream": True
                }

            else:
                resp = requests.post(
                    NVIDIA_API_URL,
                    headers=headers,
                    json=payload,
                    timeout=_timeout
                )
                resp.raise_for_status()

                content = resp.json()["choices"][0]["message"]["content"]

                return {
                    "success": True,
                    "content": content,
                    "error": None,
                    "attempt": attempt + 1,
                    "model": _model,
                    "stream": False
                }

        except Exception as e:
            last_error = str(e)
            time.sleep(2 * (attempt + 1))

    return {
        "success": False,
        "content": None,
        "error": last_error,
        "attempt": NVIDIA_MAX_RETRIES + 1,
        "model": _model,
        "stream": _stream
    }

print("공통 유틸리티 준비 완료")

# %% [markdown]
"""
## 5. BigKinds Excel Ingestion & DB Schema

BigKinds Excel 파일을 표준화하고 SQLite DB에 적재합니다. 날짜 파싱과 원천 컬럼 정규화를 포함합니다.

"""

# %%
# =========================================
# Cell 5. 정형 DB 스키마 및 BigKinds 적재 Tool
# =========================================
def get_conn(db_path: str = DB_PATH):
    return sqlite3.connect(db_path)


def init_db_tool(db_path: str = DB_PATH):
    conn = get_conn(db_path)
    cur = conn.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS query_logs (
        query_id TEXT PRIMARY KEY, user_query TEXT,
        query_plan_json TEXT, created_at TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS agent_outputs (
        output_id TEXT PRIMARY KEY, user_query TEXT, final_status TEXT,
        final_answer TEXT, output_excel_path TEXT, created_at TEXT)""")
    conn.commit()
    conn.close()


def load_bigkinds_file_tool(file_path: str) -> pd.DataFrame:
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"파일 없음: {file_path}")
    if file_path.lower().endswith((".xlsx",".xls")):
        return pd.read_excel(file_path)
    try:
        return pd.read_csv(file_path, encoding="utf-8-sig")
    except UnicodeDecodeError:
        return pd.read_csv(file_path, encoding="cp949")

def parse_bigkinds_date(value):
    """
    빅카인즈 일자 컬럼 파싱 함수.
    20260518, '20260518', '2026-05-18', datetime 모두 안전하게 처리한다.
    """

    if pd.isna(value):
        return None

    # 이미 datetime 형태인 경우
    if isinstance(value, (pd.Timestamp, datetime)):
        return pd.to_datetime(value).strftime("%Y-%m-%d")

    text = str(value).strip()

    # 엑셀에서 20260518.0처럼 들어온 경우 처리
    if text.endswith(".0"):
        text = text[:-2]

    # 20260518 형식
    if re.fullmatch(r"\d{8}", text):
        return pd.to_datetime(text, format="%Y%m%d", errors="coerce").strftime("%Y-%m-%d")

    # 2026-05-18 등 일반 날짜 형식
    parsed = pd.to_datetime(text, errors="coerce")

    if pd.isna(parsed):
        return None

    return parsed.strftime("%Y-%m-%d")


def standardize_columns_tool(df: pd.DataFrame) -> pd.DataFrame:
    col_map = {
        "뉴스 식별자":"bigkinds_news_id","뉴스식별자":"bigkinds_news_id",
        "일자":"published_date","언론사":"outlet","기고자":"author","제목":"title",
        "통합 분류1":"category_1","통합분류1":"category_1",
        "통합 분류2":"category_2","통합분류2":"category_2",
        "통합 분류3":"category_3","통합분류3":"category_3",
        "키워드":"keywords","특성추출":"features","본문":"content",
        "URL":"url","분석제외 여부":"exclude_yn",
        "인물":"persons","위치":"locations","기관":"organizations"
    }
    df = df.rename(columns={c: col_map.get(c, c) for c in df.columns}).copy()
    required = ["bigkinds_news_id","published_date","outlet","author","title",
                "category_1","category_2","category_3","keywords","features",
                "content","url","exclude_yn"]
    for col in required:
        if col not in df.columns:
            df[col] = ""
    if "published_date" in df.columns:
      df["published_date"] = df["published_date"].apply(parse_bigkinds_date)
    else:
      df["published_date"] = None
    for col in ["title","outlet","keywords","features","content","url","category_1","category_2","category_3"]:
        df[col] = df[col].apply(normalize_text)
    if "article_id" not in df.columns:
        df["article_id"] = [f"A{i+1:07d}" for i in range(len(df))]
    df["collection_source"] = "BIGKINDS"
    df["source_file"] = os.path.basename(FILE_PATH)
    df["loaded_at"]   = now_str()
    return df


def deduplicate_articles_tool(df: pd.DataFrame) -> pd.DataFrame:
    before = len(df)
    if "exclude_yn" in df.columns:
        df = df.loc[~df["exclude_yn"].astype(str).str.upper().isin(["Y","YES","1","TRUE"])].copy()
    df["dedup_key"] = df.apply(
        lambda r: r["url"] if normalize_text(r.get("url")) else f"{r.get('published_date')}|{r.get('outlet')}|{r.get('title')}",
        axis=1
    )
    df = df.drop_duplicates(subset=["dedup_key"]).drop(columns=["dedup_key"])
    print(f"중복/제외 처리: {before} → {len(df)}")
    return df.reset_index(drop=True)


def save_raw_articles_tool(df: pd.DataFrame, db_path: str = DB_PATH) -> int:
    init_db_tool(db_path)
    conn = get_conn(db_path)
    df.to_sql("raw_articles", conn, if_exists="replace", index=False)
    conn.close()
    return len(df)


def load_and_save_bigkinds_tool(file_path: str, db_path: str = DB_PATH) -> Dict[str, Any]:
    raw  = load_bigkinds_file_tool(file_path)
    std  = standardize_columns_tool(raw)
    dedup = deduplicate_articles_tool(std)
    dedup["source_file"] = os.path.basename(file_path)
    n = save_raw_articles_tool(dedup, db_path)
    return {"raw_count": len(raw), "saved_count": n, "columns": list(dedup.columns)}

print("DB/적재 Tool 준비 완료")

# %% [markdown]
"""
## 6. Domain Classification & Risk Event Tagging

Rule 기반 도메인 분류와 위험 이벤트 후보 생성을 수행합니다. PF/SPF 오탐 방지 로직을 포함합니다.

"""

# %%
# =========================================
# Cell 6. 도메인 분류 / 위험 이벤트 태깅 Tool
# =========================================

def get_latest_date_tool(db_path: str = DB_PATH) -> str:
    """
    raw_articles 테이블에서 최신 기사 일자를 조회한다.
    """
    conn = get_conn(db_path)
    try:
        row = pd.read_sql(
            "SELECT MAX(published_date) AS d FROM raw_articles",
            conn
        ).iloc[0]
        return safe_str(row["d"])
    except Exception:
        return ""
    finally:
        conn.close()


def fetch_raw_articles_tool(db_path: str = DB_PATH, date_scope: str = "latest") -> List[Dict[str, Any]]:
    """
    분석 대상 원천 기사를 DB에서 조회한다.

    date_scope:
    - latest: DB 내 최신일자
    - all / 전체 / "": 전체 기사
    - YYYY-MM-DD: 해당 일자 기사
    """
    conn = get_conn(db_path)
    try:
        if date_scope == "latest":
            latest = get_latest_date_tool(db_path)
            if not latest:
                return []
            df = pd.read_sql(
                "SELECT * FROM raw_articles WHERE published_date = ?",
                conn,
                params=[latest]
            )

        elif date_scope in ("all", "전체", ""):
            df = pd.read_sql("SELECT * FROM raw_articles", conn)

        else:
            df = pd.read_sql(
                "SELECT * FROM raw_articles WHERE published_date = ?",
                conn,
                params=[date_scope]
            )

        return df.to_dict("records")

    finally:
        conn.close()


def keyword_exists(text: str, keyword: str) -> bool:
    """
    위험 이벤트 및 도메인 키워드 매칭 함수.

    특히 PF는 SPF, APF 등 단어 내부에 포함된 경우를 오탐하지 않도록 처리한다.
    예:
    - '부동산 PF' → 매칭
    - 'PF 대출' → 매칭
    - 'SPF 선크림' → 매칭 제외
    """

    text = str(text or "")
    keyword = str(keyword or "").strip()

    if not keyword:
        return False

    # PF 오탐 방지
    if keyword.upper() == "PF":
        pf_patterns = [
            r"(?<![A-Za-z])PF(?![A-Za-z])",
            r"부동산\s*PF",
            r"프로젝트\s*파이낸싱",
            r"프로젝트파이낸싱",
            r"프로젝트금융",
            r"PF\s*수수료",
            r"PF\s*만기",
            r"PF\s*대출",
            r"PF\s*사업",
            r"PF\s*부실",
            r"PF\s*시장",
            r"PF\s*리스크",
        ]
        return any(re.search(pattern, text, flags=re.IGNORECASE) for pattern in pf_patterns)

    return keyword in text


def extract_evidence_sentence_tool(text: str, matched_keywords: List[str], max_len: int = 260) -> str:
    """
    매칭 키워드가 포함된 근거 문장을 추출한다.
    복잡한 look-behind 정규식은 사용하지 않고 안전하게 문장 후보를 분리한다.
    """

    text = safe_str(text).replace("\n", " ").replace("\r", " ")
    matched_keywords = matched_keywords or []

    if not text:
        return ""

    # 간단하고 안전한 문장 분리
    sentence_candidates = re.split(r"[.!?。]|다\s|요\s|[\n\r]+", text)

    for sent in sentence_candidates:
        sent = sent.strip()
        if not sent:
            continue

        if any(kw in sent for kw in matched_keywords):
            return sent[:max_len]

    # 문장 단위로 못 찾으면 전체 텍스트 앞부분 반환
    return text[:max_len]


def domain_keyword_match_tool(article: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    기사 1건에 대해 4개 도메인 관련성을 Rule 기반으로 탐지한다.

    도메인:
    - 건설경기
    - 금융시장
    - 건설자재
    - 노동시장
    """

    text = make_article_text(article)
    results = []

    for domain, kws in DOMAIN_KEYWORDS.items():
        matched = [kw for kw in kws if keyword_exists(text, kw)]

        if not matched:
            continue

        matched_unique = sorted(set(matched), key=matched.index)

        results.append({
            "article_id": article.get("article_id"),
            "domain": domain,
            "matched_keywords": ", ".join(matched_unique),
            "domain_score": len(matched_unique),
            "created_at": now_str()
        })

    return results


def risk_event_dictionary_tool(
    article: Dict[str, Any],
    domain_tags: Optional[List[Dict[str, Any]]] = None
) -> List[Dict[str, Any]]:
    """
    기사 1건에 대해 위험 이벤트를 Rule 기반으로 탐지한다.

    주의:
    - RISK_EVENTS.items() 반복 변수명은 event, kws로 고정
    - matched 계산에서도 반드시 kws를 사용
    - keyword_exists()를 사용해 PF 오탐을 방지
    """

    text = make_article_text(article)
    domain_tags = domain_tags or []

    article_domains = [
        safe_str(d.get("domain"))
        for d in domain_tags
        if d.get("domain")
    ]

    results = []

    for event, kws in RISK_EVENTS.items():
        matched = [kw for kw in kws if keyword_exists(text, kw)]

        if not matched:
            continue

        matched_unique = sorted(set(matched), key=matched.index)

        score = len(matched_unique)

        # 해당 도메인의 주요 이벤트이면 가중치 부여
        if any(event in DOMAIN_EVENT_HINTS.get(domain, []) for domain in article_domains):
            score += 2

        if score >= 4:
            level = "높음"
        elif score >= 2:
            level = "중간"
        else:
            level = "낮음"

        results.append({
            "event_result_id": str(uuid.uuid4()),
            "article_id": article.get("article_id"),
            "published_date": article.get("published_date"),
            "outlet": article.get("outlet"),
            "title": article.get("title"),
            "url": article.get("url"),
            "domains_rule": ", ".join(sorted(set(article_domains))),
            "event_type": event,
            "matched_keywords": ", ".join(matched_unique),
            "evidence_text": extract_evidence_sentence_tool(text, matched_unique),
            "risk_score_rule": score,
            "risk_level_rule": level,
            "detected_by": "Rule",
            "created_at": now_str()
        })

    return results


def save_domain_tags_tool(tags: List[Dict[str, Any]], db_path: str = DB_PATH) -> int:
    """
    도메인 분류 결과를 article_domain_tags 테이블에 저장한다.
    """
    conn = get_conn(db_path)

    try:
        if tags:
            df = pd.DataFrame(tags)
        else:
            df = pd.DataFrame(columns=[
                "article_id",
                "domain",
                "matched_keywords",
                "domain_score",
                "created_at"
            ])

        df.to_sql("article_domain_tags", conn, if_exists="replace", index=False)
        return len(df)

    finally:
        conn.close()


def save_risk_events_tool(events: List[Dict[str, Any]], db_path: str = DB_PATH) -> int:
    """
    위험 이벤트 탐지 결과를 risk_event_candidates 테이블에 저장한다.
    """
    conn = get_conn(db_path)

    try:
        if events:
            df = pd.DataFrame(events)
        else:
            df = pd.DataFrame(columns=[
                "event_result_id",
                "article_id",
                "published_date",
                "outlet",
                "title",
                "url",
                "domains_rule",
                "event_type",
                "matched_keywords",
                "evidence_text",
                "risk_score_rule",
                "risk_level_rule",
                "detected_by",
                "created_at"
            ])

        df.to_sql("risk_event_candidates", conn, if_exists="replace", index=False)
        return len(df)

    finally:
        conn.close()


def build_event_summary_tool(events: List[Dict[str, Any]], db_path: str = DB_PATH) -> pd.DataFrame:
    """
    위험 이벤트 결과를 일자·이벤트 유형별로 집계한다.
    """
    df = pd.DataFrame(events)

    if df.empty:
        summary = pd.DataFrame(columns=[
            "published_date",
            "event_type",
            "article_count",
            "outlet_count",
            "max_rule_score",
            "risk_level_summary"
        ])

    else:
        summary = (
            df.groupby(["published_date", "event_type"], dropna=False)
            .agg(
                article_count=("article_id", "nunique"),
                outlet_count=("outlet", "nunique"),
                max_rule_score=("risk_score_rule", "max")
            )
            .reset_index()
        )

        def summarize_level(row):
            if row["outlet_count"] >= 3 or row["max_rule_score"] >= 4:
                return "높음"
            if row["outlet_count"] >= 2 or row["max_rule_score"] >= 2:
                return "중간"
            return "낮음"

        summary["risk_level_summary"] = summary.apply(summarize_level, axis=1)

    conn = get_conn(db_path)

    try:
        summary.to_sql("event_summary", conn, if_exists="replace", index=False)
        return summary

    finally:
        conn.close()


def build_rag_documents_tool(db_path: str = DB_PATH) -> int:
    """
    risk_event_candidates 결과를 기반으로 RAG 문서 테이블을 생성한다.

    현재 구조:
    - 위험 이벤트 row 단위로 RAG 문서 생성
    - 동일 article_id가 여러 이벤트로 중복 저장될 수 있음
    - Vector Search 출력 단계에서 article_id 기준 중복 제거 권장
    """
    conn = get_conn(db_path)

    try:
        risk_df = pd.read_sql("SELECT * FROM risk_event_candidates", conn)

        if risk_df.empty:
            empty_df = pd.DataFrame(columns=[
                "doc_id",
                "article_id",
                "document_text",
                "metadata_json",
                "created_at"
            ])

            empty_df.to_sql("rag_documents", conn, if_exists="replace", index=False)
            return 0

        docs = []

        for _, row in risk_df.iterrows():
            title = safe_str(row.get("title", ""))
            outlet = safe_str(row.get("outlet", ""))
            published_date = safe_str(row.get("published_date", ""))
            domains_rule = safe_str(row.get("domains_rule", ""))
            event_type = safe_str(row.get("event_type", ""))
            matched_keywords = safe_str(row.get("matched_keywords", ""))
            evidence_text = safe_str(row.get("evidence_text", ""))
            url = safe_str(row.get("url", ""))

            document_text = (
                f"[기사 제목] {title}\n"
                f"[언론사] {outlet}\n"
                f"[일자] {published_date}\n"
                f"[도메인] {domains_rule}\n"
                f"[위험 이벤트] {event_type}\n"
                f"[매칭 키워드] {matched_keywords}\n"
                f"[판단 근거] {evidence_text}\n"
                f"[URL] {url}"
            ).strip()

            metadata = {
                "article_id": safe_str(row.get("article_id", "")),
                "published_date": published_date,
                "outlet": outlet,
                "title": title,
                "domains_rule": domains_rule,
                "event_type": event_type,
                "risk_level_rule": safe_str(row.get("risk_level_rule", "")),
                "matched_keywords": matched_keywords,
                "url": url
            }

            docs.append({
                "doc_id": str(uuid.uuid4()),
                "article_id": safe_str(row.get("article_id", "")),
                "document_text": document_text,
                "metadata_json": json.dumps(metadata, ensure_ascii=False),
                "created_at": now_str()
            })

        pd.DataFrame(docs).to_sql("rag_documents", conn, if_exists="replace", index=False)
        return len(docs)

    finally:
        conn.close()


def default_monitoring_tool(db_path: str = DB_PATH, date_scope: str = "latest") -> Dict[str, Any]:
    """
    사용자 질의 없이 기본적으로 수행되는 일일 모니터링 기능.

    수행 작업:
    1. 분석 대상 기사 조회
    2. 도메인 분류
    3. 위험 이벤트 태깅
    4. event_summary 저장
    5. rag_documents 저장
    """
    articles = fetch_raw_articles_tool(db_path, date_scope=date_scope)

    domain_tags = []
    risk_events = []

    for article in tqdm(articles, desc="도메인/위험 이벤트 분석"):
        tags = domain_keyword_match_tool(article)
        domain_tags.extend(tags)

        events = risk_event_dictionary_tool(article, tags)
        risk_events.extend(events)

    domain_count = save_domain_tags_tool(domain_tags, db_path)
    event_count = save_risk_events_tool(risk_events, db_path)
    summary_df = build_event_summary_tool(risk_events, db_path)
    rag_count = build_rag_documents_tool(db_path)

    return {
        "analyzed_article_count": len(articles),
        "domain_tag_count": domain_count,
        "risk_event_count": event_count,
        "event_summary_count": len(summary_df),
        "rag_document_count": rag_count
    }


print("Cell 6 준비 완료: 도메인 분류 / 위험 이벤트 태깅 Tool")

# %% [markdown]
"""
## 7. Rule Query Planner & Tool Executor

자연어 질의를 Rule 기반 실행 계획으로 바꾸고, 후보 기사 테이블을 조회합니다.

"""

# %%
# =========================================
# Cell 7. Query Planner (Rule 전용) 및 Tool Executor
# ★ v4: USE_LLM_QUERY_PLANNER=False 기준으로 Rule Planner가 기본 동작
# =========================================

def default_query_plan() -> Dict[str, Any]:
    return {
        "date_scope":       "latest",
        "primary_domains":  ["건설경기","금융시장","건설자재","노동시장"],
        "secondary_domains": [],
        "risk_events":      [],
        "keywords":         [],
        "task_type":        "daily_monitoring",
        "need_llm_judgement": True,
        "top_n":            LLM_FINAL_TOP_N,
        "output_focus":     "briefing",
        "planner_status":   "default"
    }


def rule_query_planner_tool(user_query: str) -> Dict[str, Any]:
    """자연어 질의를 키워드 매칭으로 즉시 실행 계획으로 변환. LLM 불필요."""
    q = normalize_text(user_query)
    plan = default_query_plan()
    plan["planner_status"] = "rule_only"

    # 도메인 추출
    domains = [d for d in DOMAIN_KEYWORDS.keys() if d in q]
    if any(w in q for w in ["전체","종합","모든","모두"]):
        plan["primary_domains"] = list(DOMAIN_KEYWORDS.keys())
    elif domains:
        plan["primary_domains"]   = [domains[0]]
        plan["secondary_domains"] = domains[1:]

    # 날짜 범위
    if any(w in q for w in ["오늘","금일","최신","최근"]):
        plan["date_scope"] = "latest"
    elif "전체 기간" in q or "전체기간" in q:
        plan["date_scope"] = "all"

    # 위험 이벤트 힌트
    plan["risk_events"] = [
        e for e, kws in RISK_EVENTS.items()
        if e in q or any(kw in q for kw in kws)
    ]

    # Task 유형
    if any(w in q for w in ["위험","선별","후속","검토","리스크"]):
        plan["task_type"] = "risk_selection"
    elif any(w in q for w in ["요약","종합","브리핑"]):
        plan["task_type"] = "event_summary"
    else:
        plan["task_type"] = "article_search"

    plan["need_llm_judgement"] = True
    return plan


def get_table_count_tool(db_path: str, table_name: str) -> int:
    conn = get_conn(db_path)
    try:
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table_name,))
        if cur.fetchone() is None:
            return 0
        cur.execute(f"SELECT COUNT(*) FROM {table_name}")
        return int(cur.fetchone()[0])
    finally:
        conn.close()


def load_risk_candidates_tool(db_path: str = DB_PATH) -> pd.DataFrame:
    if get_table_count_tool(db_path, "risk_event_candidates") == 0:
        return pd.DataFrame()
    conn = get_conn(db_path)
    try:
        return pd.read_sql("SELECT * FROM risk_event_candidates", conn)
    finally:
        conn.close()


def execute_query_plan_tool(plan: Dict[str, Any], db_path: str = DB_PATH) -> Dict[str, Any]:
    if get_table_count_tool(db_path, "risk_event_candidates") == 0:
        default_monitoring_tool(db_path, date_scope="latest")

    df    = load_risk_candidates_tool(db_path)
    debug = {"initial_candidates": len(df)}
    if df.empty:
        return {"status": "no_candidates", "result_df": df, "debug_counts": debug}

    # 날짜 필터
    date_scope = plan.get("date_scope","latest")
    if date_scope == "latest":
        latest = get_latest_date_tool(db_path)
        df = df[df["published_date"].astype(str) == latest]
    elif date_scope not in ("all","전체",""):
        df = df[df["published_date"].astype(str) == str(date_scope)]
    debug["after_date_filter"] = len(df)
    original_after_date = df.copy()

    # 도메인 필터
    target_domains = list(dict.fromkeys(
        (plan.get("primary_domains") or []) + (plan.get("secondary_domains") or [])
    ))
    if target_domains:
        pat = "|".join(re.escape(d) for d in target_domains)
        df = df[df["domains_rule"].astype(str).str.contains(pat, na=False)]
    debug["after_domain_filter"] = len(df)

    # 이벤트 필터 (0건이면 완화)
    risk_events = plan.get("risk_events") or []
    if risk_events:
        epat = "|".join(re.escape(e) for e in risk_events)
        ef   = df[df["event_type"].astype(str).str.contains(epat, na=False)]
        debug["after_event_filter_strict"] = len(ef)
        if len(ef) > 0:
            df = ef
        elif len(df) == 0:
            df = original_after_date.copy()
            debug["relaxation"] = "domain_filter_relaxed"
        else:
            debug["relaxation"] = "event_filter_relaxed"

    # 키워드 필터 (0건이면 완화)
    keywords = plan.get("keywords") or []
    if keywords:
        kpat = "|".join(re.escape(k) for k in keywords)
        kf = df[
            df["title"].astype(str).str.contains(kpat, na=False) |
            df["matched_keywords"].astype(str).str.contains(kpat, na=False) |
            df["evidence_text"].astype(str).str.contains(kpat, na=False)
        ]
        debug["after_keyword_filter_strict"] = len(kf)
        df = kf if len(kf) > 0 else df

    # 정렬 및 상위 N
    top_n = int(plan.get("top_n") or LLM_FINAL_TOP_N)
    if "risk_score_rule" in df.columns:
        df["risk_score_rule"] = pd.to_numeric(df["risk_score_rule"], errors="coerce").fillna(0)
        df = df.sort_values(["risk_score_rule","event_type"], ascending=[False, True])
    df = df.drop_duplicates(subset=["article_id","event_type"]).head(top_n).reset_index(drop=True)
    debug["final_candidate_count"] = len(df)

    return {"status": "tool_completed" if len(df) > 0 else "no_candidates", "result_df": df, "debug_counts": debug}

print("Query Planner / Tool Executor 준비 완료")

# %% [markdown]
"""
## 8. LLM Final Judgement

Rule 후보 기사에 대해 NVIDIA LLM 최종 판단을 수행합니다. 전수 판단이 아닌 상위 후보 검토 보조 모듈입니다.

"""

# %%
# =========================================
# Cell 8. NVIDIA Batch Final Risk Judgement Tool
# =========================================

FINAL_RISK_JUDGEMENT_SYSTEM_PROMPT = """
너는 건설산업 조기경보시스템(EWS)의 뉴스 위험 판단 에이전트다.

너의 역할은 Rule 기반 필터링을 거친 후보 기사 목록을 검토하여,
각 기사가 건설경기, 금융시장, 건설자재, 노동시장 위험요인으로 작용할 가능성이 있는지 최종 판단하는 것이다.

반드시 제공된 기사 제목, 언론사, 일자, 도메인 후보, 위험 이벤트 후보, 키워드, 근거 텍스트, URL만 근거로 판단한다.
존재하지 않는 기사, 수치, 출처, URL을 생성하지 않는다.

출력은 반드시 JSON 배열로만 작성한다.
마크다운 코드블록, 설명문, 주석을 붙이지 않는다.

각 원소는 다음 필드를 포함한다.

- article_id
- is_risk_relevant: true 또는 false
- related_domains: ["건설경기", "금융시장", "건설자재", "노동시장"] 중 하나 이상
- risk_event_type
- risk_level: "높음" | "중간" | "낮음"
- judgement_reason: 기사 내용에 근거한 1문장 판단 사유
- followup_required: true 또는 false
"""


def chunk_list(items, chunk_size):
    for i in range(0, len(items), chunk_size):
        yield items[i:i + chunk_size]


def dataframe_to_candidate_articles(candidate_df):
    """
    Tool Executor 결과 DataFrame을 LLM 판단용 list[dict]로 변환한다.
    """
    if candidate_df is None:
        return []

    if isinstance(candidate_df, list):
        return candidate_df

    if not isinstance(candidate_df, pd.DataFrame):
        return []

    if candidate_df.empty:
        return []

    rows = []

    for _, row in candidate_df.iterrows():
        rows.append({
            "article_id": row.get("article_id", ""),
            "title": row.get("title", ""),
            "outlet": row.get("outlet", ""),
            "published_date": row.get("published_date", ""),
            "domains_rule": row.get("domains", row.get("domains_rule", "")),
            "event_type": row.get("event_type", ""),
            "matched_keywords": row.get("matched_keywords", ""),
            "evidence_text": row.get("evidence_text", row.get("content_excerpt", "")),
            "url": row.get("url", "")
        })

    return rows


def nvidia_batch_final_judgement_tool(candidate_articles, user_query, top_n=None, batch_size=None):
    """
    Rule 기반 후보 기사를 NVIDIA LLM이 최종 판단한다.

    개선 사항:
    - 후보 기사 전체를 한 번에 보내지 않고 batch_size 단위로 분할
    - 빠른 모델 사용
    - streaming 호출 지원
    - 일부 batch 성공 시 partial success 처리
    - LLM 실패 시 Rule 결과를 최종 판단으로 둔갑시키지 않음
    """

    if top_n is None:
        top_n = LLM_FINAL_TOP_N

    if batch_size is None:
        batch_size = LLM_BATCH_SIZE

    candidate_articles = candidate_articles or []
    candidates = candidate_articles[:top_n]

    all_judgements = []
    batch_errors = []

    if not candidates:
        return {
            "status": "no_candidates",
            "judgements": [],
            "errors": [],
            "candidate_count": 0,
            "completed_count": 0
        }

    total_batches = -(-len(candidates) // batch_size)

    for batch_idx, batch in enumerate(chunk_list(candidates, batch_size), start=1):
        article_blocks = []

        for i, art in enumerate(batch, start=1):
            evidence = str(art.get("evidence_text", "") or "")
            evidence = evidence.replace("\n", " ").replace("\r", " ")
            evidence = evidence[:LLM_CONTENT_MAX_CHARS]

            article_blocks.append(f"""
[기사 {i}]
article_id: {art.get('article_id', '')}
title: {art.get('title', '')}
outlet: {art.get('outlet', '')}
published_date: {art.get('published_date', '')}
domains_rule: {art.get('domains_rule', '')}
event_type_rule: {art.get('event_type', '')}
matched_keywords: {art.get('matched_keywords', '')}
evidence_text: {evidence}
url: {art.get('url', '')}
""")

        user_prompt = (
            f"[사용자 질의]\n{user_query}\n\n"
            f"[Rule 기반 후보 기사 목록 — batch {batch_idx}/{total_batches}]\n"
            + "\n".join(article_blocks)
            + "\n\n위 후보 기사들을 검토하여 사용자 질의에 부합하는 위험 기사인지 최종 판단해라.\n"
            + "반드시 JSON 배열로만 출력하라.\n"
            + "설명 문단을 길게 쓰지 말고, judgement_reason은 기사당 1문장으로 제한하라.\n"
        )

        messages = [
            {
                "role": "system",
                "content": FINAL_RISK_JUDGEMENT_SYSTEM_PROMPT
            },
            {
                "role": "user",
                "content": user_prompt
            }
        ]

        result = call_nvidia_llm_safe(
            messages,
            max_tokens=700,
            timeout_sec=NVIDIA_TIMEOUT_SEC,
            model=NVIDIA_MODEL_FAST,
            stream=True
        )

        if not result.get("success"):
            batch_errors.append({
                "batch_idx": batch_idx,
                "error": result.get("error", "unknown error")
            })
            continue

        parsed = safe_json_loads(result.get("content", ""), default=[])

        if isinstance(parsed, list):
            all_judgements.extend(parsed)
        elif isinstance(parsed, dict):
            all_judgements.append(parsed)
        else:
            batch_errors.append({
                "batch_idx": batch_idx,
                "error": "JSON 파싱 실패",
                "raw_response": str(result.get("content", ""))[:500]
            })

    if all_judgements and batch_errors:
        status = "llm_completed_with_partial_errors"
    elif all_judgements:
        status = "llm_completed"
    else:
        status = "llm_failed"

    return {
        "status": status,
        "judgements": all_judgements,
        "errors": batch_errors,
        "candidate_count": len(candidates),
        "completed_count": len(all_judgements)
    }


def build_final_answer_tool(user_query, query_plan, tool_result, judgement_result):
    """
    Tool 결과와 LLM 최종 판단 결과를 기반으로 최종 답변을 생성한다.
    Rule 결과를 LLM 최종 판단으로 둔갑시키지 않는다.
    """

    status = judgement_result.get("status", "unknown")
    judgements = judgement_result.get("judgements", [])
    errors = judgement_result.get("errors", [])

    selected = [
        item for item in judgements
        if item.get("is_risk_relevant") is True
    ]

    if status in ["llm_completed", "llm_completed_with_partial_errors"]:
        lines = []
        lines.append("[LLM 최종 판단 완료]" if status == "llm_completed" else "[LLM 최종 판단 부분 완료]")
        lines.append("")
        lines.append(f"1. 사용자 질의: {user_query}")
        tool_candidate_count = (
    tool_result.get("final_candidate_count")
    or tool_result.get("debug", {}).get("final_candidate_count")
    or 0
)

        lines.append(f"2. Tool 기반 후보 기사 수: {tool_candidate_count}건")
        lines.append(f"3. LLM 검토 기사 수: {judgement_result.get('candidate_count', 0)}건")
        lines.append(f"4. LLM이 위험요인 가능 기사로 판단한 기사 수: {len(selected)}건")

        if errors:
            lines.append("")
            lines.append("일부 batch에서 오류가 발생했으나, 성공한 batch의 LLM 판단 결과는 반영했다.")
            lines.append(f"오류 batch 수: {len(errors)}")

        lines.append("")
        lines.append("주요 선별 기사:")

        if selected:
            for idx, item in enumerate(selected[:10], start=1):
                lines.append(
                    f"{idx}. article_id={item.get('article_id')} | "
                    f"위험수준={item.get('risk_level')} | "
                    f"이벤트={item.get('risk_event_type')} | "
                    f"사유={item.get('judgement_reason')}"
                )
        else:
            lines.append("- LLM이 최종적으로 위험요인 가능성이 높다고 판단한 기사는 없었다.")

        return "\n".join(lines)

    if status == "no_candidates":
        return (
            "[후보 기사 없음]\n"
            "Tool 기반 조회 결과, LLM 최종 판단에 넘길 후보 기사가 없습니다.\n"
            "질의 조건을 완화하거나 도메인·위험 이벤트 조건을 조정해보세요."
        )

    return (
        "[LLM 최종 판단 미완료]\n"
        "NVIDIA API 응답 지연 또는 응답 형식 오류로 인해 최종 위험 판단은 완료되지 않았습니다.\n"
        "아래 결과는 Tool 기반 후보 기사 목록이며, 최종 판단 결과가 아닙니다.\n"
        f"Tool 후보 기사 수: {tool_result.get('final_candidate_count', 0)}건\n"
        f"오류: {errors if errors else judgement_result.get('error', 'unknown')}"
    )

# %% [markdown]
"""
## 9. Report Export

Agent 결과를 Excel 리포트로 저장합니다.

"""

# %%
# =========================================
# Cell 9. 보고서 저장 Tool
# =========================================
def export_excel_report_tool(db_path: str = DB_PATH, output_path: str = OUTPUT_EXCEL_PATH,
                              candidate_df: Optional[pd.DataFrame] = None) -> str:
    conn = get_conn(db_path)
    try:
        tables = ["raw_articles","article_domain_tags","risk_event_candidates",
                  "event_summary","rag_documents","llm_final_judgements","query_logs","agent_outputs"]
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for table in tables:
                try:
                    if get_table_count_tool(db_path, table) > 0:
                        df = pd.read_sql(f"SELECT * FROM {table}", conn)
                        df.head(5000).to_excel(writer, sheet_name=table[:31], index=False)
                except Exception:
                    continue
            if candidate_df is not None and not candidate_df.empty:
                candidate_df.to_excel(writer, sheet_name="query_candidates", index=False)
        return output_path
    finally:
        conn.close()


def save_query_and_output_tool(user_query, plan, final_answer, status, output_excel_path, db_path=DB_PATH):
    init_db_tool(db_path)
    conn = get_conn(db_path)
    try:
        pd.DataFrame([{
            "query_id": str(uuid.uuid4()), "user_query": user_query,
            "query_plan_json": json.dumps(plan, ensure_ascii=False), "created_at": now_str()
        }]).to_sql("query_logs", conn, if_exists="append", index=False)
        pd.DataFrame([{
            "output_id": str(uuid.uuid4()), "user_query": user_query, "final_status": status,
            "final_answer": final_answer, "output_excel_path": output_excel_path, "created_at": now_str()
        }]).to_sql("agent_outputs", conn, if_exists="append", index=False)
    finally:
        conn.close()

print("보고서 저장 Tool 준비 완료")

# %% [markdown]
"""
## 10. LangGraph Agent Workflow

Ingestion → Monitoring → Query Planner → Tool Executor → LLM Judgement → Report 흐름을 LangGraph로 구성합니다.

"""

# %%
# =========================================
# Cell 10. LangGraph Agent State 및 Node 정의
# ★ v4: query_planner_agent_node — Rule 전용 (LLM 호출 없음)
# =========================================
from langgraph.graph import StateGraph, START, END


class EWSAgentState(TypedDict, total=False):
    source_file_path:   str
    db_path:            str
    output_excel_path:  str
    user_query:         str
    ingestion_result:   Dict[str, Any]
    monitoring_result:  Dict[str, Any]
    query_plan:         Dict[str, Any]
    tool_result:        Dict[str, Any]
    judgement_result:   Dict[str, Any]
    final_answer:       str
    final_status:       str
    error:              str


def ingestion_agent_node(state: EWSAgentState) -> EWSAgentState:
    db_path   = state.get("db_path", DB_PATH)
    file_path = state.get("source_file_path", "")
    if file_path:
        state["ingestion_result"] = load_and_save_bigkinds_tool(file_path, db_path)
        print("✓ Ingestion:", state["ingestion_result"])
    return state


def default_monitoring_agent_node(state: EWSAgentState) -> EWSAgentState:
    db_path = state.get("db_path", DB_PATH)
    if get_table_count_tool(db_path, "raw_articles") == 0:
        print("raw_articles 없음 — 모니터링 건너뜀")
        return state
    state["monitoring_result"] = default_monitoring_tool(db_path, date_scope="latest")
    print("✓ Monitoring:", state["monitoring_result"])
    return state


def query_planner_agent_node(state: EWSAgentState) -> EWSAgentState:
    """★ v4: USE_LLM_QUERY_PLANNER=False → Rule Planner만 실행 (즉시 완료)."""
    user_query = state.get("user_query", "")
    if not user_query:
        state["query_plan"] = default_query_plan()
        return state
    # USE_LLM_QUERY_PLANNER 플래그가 True로 변경된 경우에도 안전하게 동작
    if USE_LLM_QUERY_PLANNER:
        from copy import deepcopy
        plan = rule_query_planner_tool(user_query)  # LLM Planner 미구현 시 Rule 사용
        plan["planner_status"] = "rule_fallback_llm_not_implemented"
    else:
        plan = rule_query_planner_tool(user_query)
    state["query_plan"] = plan
    print("✓ Query Plan:", json.dumps(plan, ensure_ascii=False))
    return state


def tool_executor_agent_node(state: EWSAgentState) -> EWSAgentState:
    db_path = state.get("db_path", DB_PATH)
    plan    = state.get("query_plan", default_query_plan())
    result  = execute_query_plan_tool(plan, db_path)
    state["tool_result"] = result
    print(f"✓ Tool Executor: {result.get('status')} | debug={result.get('debug_counts')}")
    return state


def final_judgement_agent_node(state: EWSAgentState) -> EWSAgentState:
    """★ v4 핵심: LLM이 최종 판단을 수행하는 유일한 노드."""
    user_query   = state.get("user_query", "기본 일일 모니터링")
    plan         = state.get("query_plan", default_query_plan())
    tool_result  = state.get("tool_result", {})
    candidate_df = tool_result.get("result_df", pd.DataFrame())

    if USE_LLM_FINAL_JUDGEMENT:
        candidates = dataframe_to_candidate_articles(candidate_df) if isinstance(candidate_df, pd.DataFrame) else []
        judgement = nvidia_batch_final_judgement_tool(
    candidate_articles=candidates,
    user_query=user_query,
    top_n=int(state.get("llm_final_top_n", plan.get("top_n") or LLM_FINAL_TOP_N)),
    batch_size=int(state.get("llm_batch_size", LLM_BATCH_SIZE))
)
    else:
        judgement = {"status":"llm_disabled","judgements":[],"candidate_count":len(candidate_df)}

    state["judgement_result"] = judgement
    save_final_judgements_tool(judgement, state.get("db_path", DB_PATH))
    print(f"✓ Final Judgement: {judgement.get('status')} | 후보={judgement.get('candidate_count')} | 완료={judgement.get('completed_count')}")
    return state


def final_answer_agent_node(state: EWSAgentState) -> EWSAgentState:
    db_path     = state.get("db_path", DB_PATH)
    output_path = state.get("output_excel_path", OUTPUT_EXCEL_PATH)
    user_query  = state.get("user_query", "기본 일일 모니터링")
    plan        = state.get("query_plan", default_query_plan())
    tool_result = state.get("tool_result", {})
    judgement   = state.get("judgement_result", {})

    final_answer = build_final_answer_tool(user_query, plan, tool_result, judgement)
    candidate_df = tool_result.get("result_df", pd.DataFrame())
    report_path  = export_excel_report_tool(db_path, output_path, candidate_df)
    status       = judgement.get("status","unknown")
    save_query_and_output_tool(user_query, plan, final_answer, status, report_path, db_path)

    state["final_answer"]       = final_answer
    state["final_status"]       = status
    state["output_excel_path"]  = report_path
    print(f"✓ Final Answer 완료 → {report_path}")
    return state


def should_ingest(state: EWSAgentState) -> str:
    return "ingest" if state.get("source_file_path") else "plan"

def after_monitoring_route(state: EWSAgentState) -> str:
    return "plan"

# ── Graph 구성 ──────────────────────────────────────────────────────────
graph = StateGraph(EWSAgentState)
graph.add_node("ingestion",         ingestion_agent_node)
graph.add_node("default_monitoring", default_monitoring_agent_node)
graph.add_node("query_planner",     query_planner_agent_node)
graph.add_node("tool_executor",     tool_executor_agent_node)
graph.add_node("final_judgement",   final_judgement_agent_node)
graph.add_node("final_answer",      final_answer_agent_node)

graph.add_conditional_edges(START, should_ingest, {"ingest":"ingestion","plan":"query_planner"})
graph.add_edge("ingestion",          "default_monitoring")
graph.add_edge("default_monitoring", "query_planner")
graph.add_edge("query_planner",      "tool_executor")
graph.add_edge("tool_executor",      "final_judgement")
graph.add_edge("final_judgement",    "final_answer")
graph.add_edge("final_answer",        END)

ews_graph = graph.compile()
print("✓ LangGraph EWS Multi-Agent 구성 완료")

# %% [markdown]
"""
## 11. Daily Monitoring Run

BigKinds 파일을 적재하고 기본 일일 모니터링을 실행합니다.

"""

# %%
# =========================================
# Cell 11. 기본 일일 모니터링 실행
# ★ v4: Query Planner LLM 제거로 블로킹 없음
#        LLM 호출은 Final Judgement 배치 2회 × 30초 = 최대 ~60초
# =========================================
# ── 사전 점검 ─────────────────────────────────────────────────────────────
import os
if not os.path.exists(FILE_PATH):
    print(f"[ERROR] 파일 없음: {FILE_PATH}")
    print("Colab 좌측 파일 패널에서 BigKinds xlsx/csv 파일을 업로드한 뒤 FILE_PATH를 수정하세요.")
else:
    print(f"파일 확인: {FILE_PATH}")
    base_state = {
    "source_file_path": FILE_PATH,
    "db_path": DB_PATH,
    "output_excel_path": OUTPUT_EXCEL_PATH,

    # 기본 모니터링에서는 LLM 최종 판단을 최소 샘플만 수행
    "llm_final_top_n": 4,
    "llm_batch_size": 2
}

    base_result = ews_graph.invoke(base_state, {"recursion_limit": 30})
    print("\n" + "=" * 55)
    print("기본 모니터링 FINAL ANSWER")
    print("=" * 55)
    print(base_result.get("final_answer"))
    print("\n보고서:", base_result.get("output_excel_path"))

# %% [markdown]
"""
## 12. DB Inspection

주요 DB 테이블과 이벤트 후보 샘플을 점검합니다.

"""

# %%
# =========================================
# Cell 12. DB 상태 점검
# =========================================
conn = get_conn(DB_PATH)
try:
    tables = pd.read_sql("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name", conn)
    print("== DB 테이블 목록 ==")
    for t in tables["name"].tolist():
        try:
            cnt = pd.read_sql(f"SELECT COUNT(*) AS cnt FROM {t}", conn)["cnt"].iloc[0]
            print(f"  {t}: {cnt:,} rows")
        except Exception as e:
            print(f"  {t}: ERROR {e}")

    if get_table_count_tool(DB_PATH, "event_summary") > 0:
        print("\n[event_summary — 상위 20건]")
        display(pd.read_sql("SELECT * FROM event_summary ORDER BY article_count DESC LIMIT 20", conn))

    if get_table_count_tool(DB_PATH, "risk_event_candidates") > 0:
        print("\n[risk_event_candidates — 상위 20건]")
        display(pd.read_sql(
            "SELECT published_date,outlet,title,domains_rule,event_type,risk_score_rule,risk_level_rule "
            "FROM risk_event_candidates ORDER BY risk_score_rule DESC LIMIT 20", conn))

    if get_table_count_tool(DB_PATH, "llm_final_judgements") > 0:
        print("\n[llm_final_judgements — 상위 20건]")
        display(pd.read_sql(
            "SELECT article_id,is_risk_relevant,risk_event_type,risk_level,followup_required,"
            "judgement_reason,judgement_status FROM llm_final_judgements LIMIT 20", conn))
finally:
    conn.close()

# %% [markdown]
"""
## 13. Interactive Query Run

사용자가 자연어 질의를 입력하고 후보 기사와 LLM 판단 결과를 확인합니다.

"""

# %%
# =========================================
# Cell 13. 자연어 질의 실행 셀 (직접 입력)
# =========================================
# 예시 질의:
#   "금일 건설경기 관련 기사 중 금융시장 위험요인으로 작용할 가능성이 있는 기사를 선별해줘"
#   "오늘 기사 전체에서 건설경기·금융시장·건설자재·노동시장 위험요인을 종합해서 후속 검토 기사를 보여줘"
#   "부동산 PF와 공사비 상승 관련 위험 기사를 찾아줘"
#   "오늘 금리와 환율 변동 관련 기사 중 건설업에 영향을 줄 수 있는 기사를 선별해줘"

if get_table_count_tool(DB_PATH, "raw_articles") == 0:
    print("[ERROR] raw_articles 없음 — Cell 11을 먼저 실행하세요.")
else:
    # ── 질의 입력 ─────────────────────────────────────────────────────────
    typed_query = input("질의를 입력하세요 (Enter = 기본 질의): ").strip()
    if not typed_query:
        typed_query = "오늘 기사 전체에서 건설경기, 금융시장, 건설자재, 노동시장 위험요인을 종합해서 후속 검토가 필요한 기사만 선별해줘"
        print(f"기본 질의 사용: {typed_query}")

    interactive_state = {
    "user_query": typed_query,
    "db_path": DB_PATH,
    "output_excel_path": "/content/ews_interactive_report.xlsx",

    # 직접 질의에서는 기본 모니터링보다 조금 더 넓게 판단
    "llm_final_top_n": 6,
    "llm_batch_size": 2
}

    print("\n에이전트 실행 중...")
    interactive_result = ews_graph.invoke(interactive_state, {"recursion_limit": 30})

    # ── 출력 ──────────────────────────────────────────────────────────────
    print("\n" + "=" * 55)
    print("QUERY PLAN (Rule 기반 즉시 생성)")
    print("=" * 55)
    print(json.dumps(interactive_result.get("query_plan",{}), ensure_ascii=False, indent=2))

    print("\n" + "=" * 55)
    print("TOOL DEBUG COUNTS")
    print("=" * 55)
    print(json.dumps(interactive_result.get("tool_result",{}).get("debug_counts",{}), ensure_ascii=False, indent=2))

    print("\n" + "=" * 55)
    print("FINAL ANSWER (LLM 최종 판단)")
    print("=" * 55)
    print(interactive_result.get("final_answer"))

    print("\n보고서:", interactive_result.get("output_excel_path"))

    result_df = interactive_result.get("tool_result",{}).get("result_df", pd.DataFrame())
    if isinstance(result_df, pd.DataFrame) and not result_df.empty:
        print("\n[Rule 후보 기사 테이블]")
        display(result_df[["published_date","outlet","title","domains_rule",
                            "event_type","risk_score_rule","url"]].head(20))

# %% [markdown]
"""
## 14. Validation Report

DB 적재 규모, Rule 후보 수, LLM 판단 수, 실행 이력을 요약합니다.

"""

# %%
# =========================================
# Cell 14. EWS Agent 응답 품질 검증 리포트
# =========================================

import sqlite3
import pandas as pd
from IPython.display import display

print("=" * 55)
print("EWS Agent 응답 품질 검증 리포트")
print("=" * 55)
print("※ 본 리포트는 전체 기사 원문이 아니라 DB에 저장된 Agent 처리 결과를 기준으로 산출됩니다.")
print("※ LLM 판단 분포는 llm_final_judgements 테이블에 저장된 'LLM 최종 판단 결과'만 기준으로 합니다.")
print("※ Rule 기반 후보 수와 LLM 최종 판단 수는 서로 다를 수 있습니다.")
print("=" * 55)

conn = get_conn(DB_PATH)

try:
    # =========================================
    # 0. DB 테이블 목록 확인
    # =========================================
    table_df = pd.read_sql(
        """
        SELECT name
        FROM sqlite_master
        WHERE type='table'
        ORDER BY name
        """,
        conn
    )

    existing_tables = set(table_df["name"].tolist())

    required_tables = [
        "raw_articles",
        "article_domain_tags",
        "risk_event_candidates",
        "event_summary",
        "rag_documents",
        "llm_final_judgements",
        "query_logs",
        "agent_outputs"
    ]

    print("\n[0] DB 테이블 상태")
    for table in required_tables:
        if table in existing_tables:
            count_df = pd.read_sql(f"SELECT COUNT(*) AS n FROM {table}", conn)
            n = int(count_df["n"].iloc[0])
            print(f"  - {table}: {n:,} rows")
        else:
            print(f"  - {table}: 없음")

    # =========================================
    # 1. 전체 처리 규모 요약
    # =========================================
    raw_count = 0
    domain_count = 0
    risk_count = 0
    rag_count = 0
    llm_count = 0

    if "raw_articles" in existing_tables:
        raw_count = int(pd.read_sql("SELECT COUNT(*) AS n FROM raw_articles", conn)["n"].iloc[0])

    if "article_domain_tags" in existing_tables:
        domain_count = int(pd.read_sql("SELECT COUNT(*) AS n FROM article_domain_tags", conn)["n"].iloc[0])

    if "risk_event_candidates" in existing_tables:
        risk_count = int(pd.read_sql("SELECT COUNT(*) AS n FROM risk_event_candidates", conn)["n"].iloc[0])

    if "rag_documents" in existing_tables:
        rag_count = int(pd.read_sql("SELECT COUNT(*) AS n FROM rag_documents", conn)["n"].iloc[0])

    if "llm_final_judgements" in existing_tables:
        llm_count = int(pd.read_sql("SELECT COUNT(*) AS n FROM llm_final_judgements", conn)["n"].iloc[0])

    print("\n[1] 전체 처리 규모 요약")
    print(f"  전체 적재 기사 수(raw_articles): {raw_count:,}건")
    print(f"  도메인 태깅 결과(article_domain_tags): {domain_count:,}건")
    print(f"  Rule 기반 위험 후보(risk_event_candidates): {risk_count:,}건")
    print(f"  RAG 문서(rag_documents): {rag_count:,}건")
    print(f"  LLM 최종 판단 저장 수(llm_final_judgements): {llm_count:,}건")

    if raw_count > 0:
        print(f"  Rule 위험 후보 비율: {risk_count / raw_count * 100:.1f}%")

    if risk_count > 0:
        print(f"  LLM 최종 판단 커버리지: {llm_count / risk_count * 100:.1f}%")

    # =========================================
    # 2. LLM 최종 판단 분포
    # =========================================
    print("\n[2] LLM 최종 판단 분포")

    if "llm_final_judgements" not in existing_tables or llm_count == 0:
        print("  LLM 최종 판단 결과가 아직 없습니다.")
        print("  Cell 11 또는 Cell 13을 실행한 뒤 다시 확인하세요.")

    else:
        ldf = pd.read_sql("SELECT * FROM llm_final_judgements", conn)

        if "is_risk_relevant" in ldf.columns:
            ldf["is_risk_relevant_norm"] = ldf["is_risk_relevant"].apply(
                lambda x: True if str(x).lower() in ["1", "true", "yes"] else False
            )
        else:
            ldf["is_risk_relevant_norm"] = False

        total = len(ldf)
        related = int(ldf["is_risk_relevant_norm"].sum())
        unrelated = total - related

        followup = 0
        if "followup_required" in ldf.columns:
            followup = int(
                ldf["followup_required"].apply(
                    lambda x: True if str(x).lower() in ["1", "true", "yes"] else False
                ).sum()
            )

        print(f"  LLM 판단 총 {total:,}건")
        print(f"  위험요인 관련: {related:,}건 ({related / total * 100:.1f}%)")
        print(f"  비관련:        {unrelated:,}건")
        print(f"  후속 검토 필요: {followup:,}건")

        if "risk_level" in ldf.columns:
            print("\n  위험 수준 분포:")
            risk_level_df = (
                ldf["risk_level"]
                .fillna("미분류")
                .value_counts()
                .reset_index()
            )
            risk_level_df.columns = ["risk_level", "count"]
            display(risk_level_df)

        if "risk_event_type" in ldf.columns:
            print("\n  위험 이벤트 유형 분포(LLM):")
            event_dist = (
                ldf["risk_event_type"]
                .fillna("미분류")
                .value_counts()
                .reset_index()
            )
            event_dist.columns = ["event_type", "llm_count"]
            display(event_dist)

        if "judgement_status" in ldf.columns:
            print("\n  LLM 판단 상태:")
            print(ldf["judgement_status"].fillna("unknown").value_counts().to_dict())

        if "judgement_reason" in ldf.columns:
            empty_reason_count = int(ldf["judgement_reason"].fillna("").str.strip().eq("").sum())
            print(f"\n  LLM 판단 근거 공백: {empty_reason_count}/{total}건")

    # =========================================
    # 3. Rule ↔ LLM 비교
    # =========================================
    print("\n[3] Rule ↔ LLM 비교")

    if (
        "risk_event_candidates" in existing_tables
        and "llm_final_judgements" in existing_tables
        and risk_count > 0
        and llm_count > 0
    ):
        rdf = pd.read_sql(
            """
            SELECT
                article_id,
                event_type AS rule_event_type,
                risk_level_rule,
                risk_score_rule
            FROM risk_event_candidates
            """,
            conn
        )

        ldf = pd.read_sql(
            """
            SELECT
                article_id,
                is_risk_relevant,
                risk_event_type AS llm_event_type,
                risk_level AS llm_risk_level
            FROM llm_final_judgements
            """,
            conn
        )

        rdf_rep = (
            rdf.sort_values(["article_id", "risk_score_rule"], ascending=[True, False])
            .drop_duplicates(subset=["article_id"], keep="first")
        )

        merged = ldf.merge(rdf_rep, on="article_id", how="left")

        merged["is_risk_relevant_norm"] = merged["is_risk_relevant"].apply(
            lambda x: True if str(x).lower() in ["1", "true", "yes"] else False
        )

        comparable = merged[merged["rule_event_type"].notna()].copy()

        if comparable.empty:
            print("  Rule 후보와 LLM 판단을 article_id 기준으로 연결할 수 없습니다.")
        else:
            comparable["rule_positive"] = True
            comparable["llm_positive"] = comparable["is_risk_relevant_norm"]

            agreement = (comparable["rule_positive"] == comparable["llm_positive"]).mean() * 100

            print(f"  비교 가능 건수: {len(comparable):,}건")
            print(f"  Rule ↔ LLM 방향 일치율: {agreement:.1f}%")
            print("  ※ 표본 수가 작으면 일치율은 품질 지표로 해석하기 어렵습니다.")
            print("  ※ Rule은 후보 선별용이고, LLM은 최종 판단용이므로 일부 불일치는 정상입니다.")

            print("\n  Rule ↔ LLM 비교 샘플:")
            display(
                comparable[
                    [
                        "article_id",
                        "rule_event_type",
                        "risk_level_rule",
                        "llm_event_type",
                        "llm_risk_level",
                        "is_risk_relevant"
                    ]
                ].head(20)
            )

    else:
        print("  Rule 후보 또는 LLM 판단 결과가 부족하여 비교할 수 없습니다.")

    # =========================================
    # 4. 최근 에이전트 실행 이력
    # =========================================
    print("\n[4] 최근 에이전트 실행 이력")

    if "query_logs" in existing_tables:
        qcols_df = pd.read_sql("PRAGMA table_info(query_logs)", conn)
        qcols = qcols_df["name"].tolist()

        print(f"  query_logs 컬럼: {qcols}")

        preferred_cols = [
            "user_query",
            "query",
            "planner_status",
            "final_status",
            "status",
            "created_at",
            "executed_at",
            "timestamp"
        ]

        selected_cols = [c for c in preferred_cols if c in qcols]

        if not selected_cols:
            print("  출력 가능한 주요 컬럼이 없습니다. 전체 상위 10건을 출력합니다.")
            qdf = pd.read_sql("SELECT * FROM query_logs LIMIT 10", conn)
            display(qdf)

        else:
            order_col = None
            for c in ["created_at", "executed_at", "timestamp"]:
                if c in qcols:
                    order_col = c
                    break

            select_sql = ", ".join(selected_cols)

            if order_col:
                qdf = pd.read_sql(
                    f"""
                    SELECT {select_sql}
                    FROM query_logs
                    ORDER BY {order_col} DESC
                    LIMIT 10
                    """,
                    conn
                )
            else:
                qdf = pd.read_sql(
                    f"""
                    SELECT {select_sql}
                    FROM query_logs
                    LIMIT 10
                    """,
                    conn
                )

            if qdf.empty:
                print("  최근 실행 이력이 없습니다.")
            else:
                display(qdf)

    else:
        print("  query_logs 테이블이 없습니다.")

    # =========================================
    # 5. Rule 기반 위험 이벤트 요약
    # =========================================
    print("\n[5] 위험 이벤트 요약(Rule 기반)")

    if "event_summary" in existing_tables:
        edf = pd.read_sql(
            """
            SELECT
                event_type,
                article_count,
                outlet_count,
                max_rule_score,
                risk_level_summary
            FROM event_summary
            ORDER BY max_rule_score DESC, article_count DESC
            LIMIT 20
            """,
            conn
        )

        if edf.empty:
            print("  event_summary 결과가 없습니다.")
        else:
            display(edf)

    else:
        print("  event_summary 테이블이 없습니다.")

    print("\n검증 완료")

finally:
    conn.close()

# %% [markdown]
"""
## 15. Optional Golden Set Evaluation

수동 라벨링한 골든셋이 있을 때 LLM 판단을 간단히 검증하는 선택 셀입니다.

"""

# %%
# =========================================
# Cell 15. LLM 판단 자동 검증 (골든셋 기반 POC 평가)
# ★ v4 신규: 사용자가 직접 라벨링한 골든셋과 LLM 판단 비교
# =========================================
# 사용 방법:
#   GOLDEN_SET에 article_id와 기대 판단(expected_risk: True/False)을 직접 입력하세요.
#   Cell 11/13 실행 후 DB의 article_id를 참조하세요 (Cell 12 출력 참고).

GOLDEN_SET = [
    # {"article_id": "A0000001", "expected_risk": True,  "note": "PF 리스크 핵심 기사"},
    # {"article_id": "A0000002", "expected_risk": False, "note": "단순 부동산 시황 기사"},
    # {"article_id": "A0000003", "expected_risk": True,  "note": "금리 인상 건설업 영향"},
]

if not GOLDEN_SET:
    print("[안내] GOLDEN_SET이 비어 있습니다.")
    print("Cell 12 출력에서 article_id를 확인하고 위 GOLDEN_SET 리스트를 채워 재실행하세요.")
elif get_table_count_tool(DB_PATH, "llm_final_judgements") == 0:
    print("[ERROR] llm_final_judgements 없음 — Cell 11 또는 Cell 13을 먼저 실행하세요.")
else:
    conn = get_conn(DB_PATH)
    try:
        jdf = pd.read_sql("SELECT article_id, is_risk_relevant, risk_level, judgement_reason FROM llm_final_judgements", conn)
    finally:
        conn.close()

    results = []
    for gs in GOLDEN_SET:
        aid = gs["article_id"]
        row = jdf[jdf["article_id"] == aid]
        if row.empty:
            results.append({"article_id": aid, "expected": gs["expected_risk"],
                            "llm_judgement": "판단 없음 (후보 미포함)", "match": "N/A", "note": gs.get("note","")})
        else:
            llm_val = bool(row.iloc[0]["is_risk_relevant"])
            match   = llm_val == gs["expected_risk"]
            results.append({
                "article_id":    aid,
                "expected":      gs["expected_risk"],
                "llm_judgement": llm_val,
                "risk_level":    row.iloc[0]["risk_level"],
                "judgement_reason": row.iloc[0]["judgement_reason"][:80],
                "match":         "✓" if match else "✗",
                "note":          gs.get("note","")
            })

    rdf = pd.DataFrame(results)
    display(rdf)

    valid = rdf[rdf["match"].isin(["✓","✗"])]
    if not valid.empty:
        accuracy = (valid["match"] == "✓").mean()
        print(f"\n골든셋 정확도: {accuracy*100:.1f}% ({int((valid['match']=='✓').sum())}/{len(valid)}건 일치)")
        print("  70% 이상이면 POC 통과 수준으로 간주합니다.")

# %% [markdown]
"""
## 16. Optional Hybrid Vector Search

RAG 문서가 생성되는지 확인하기 위한 선택 셀입니다. 운영형 검색엔진이 아니라 PoC 부가 검증입니다.

"""

# %%
# =========================================
# Cell 16. Optional Hybrid Vector Search 점검
# - PoC 부가 검증용 안정화 버전
# - metadata/rule filter + vector search
# =========================================

import os
import re
import json
import time
import shutil
import pandas as pd

print("=" * 60)
print("Cell 16. Hybrid Vector Search 점검")
print("=" * 60)

# =========================================
# 0. 실행 옵션
# =========================================
CHROMA_COLLECTION_NAME = "ews_news_risk_documents"

# ChromaDB readonly/lock 오류 방지를 위해 매 실행마다 새 경로 사용
VECTOR_DB_RUNTIME_PATH = f"/content/ews_bigkinds_chroma_runtime_{int(time.time())}"

# Vector 검색 전 1차 필터링 여부
USE_HYBRID_PREFILTER = True

# 최종 출력 개수
TOP_K_DISPLAY = 7

print(f"Vector DB Runtime Path: {VECTOR_DB_RUNTIME_PATH}")
print(f"Hybrid Prefilter 사용 여부: {USE_HYBRID_PREFILTER}")

# =========================================
# 1. 라이브러리 로드
# =========================================
try:
    import chromadb
    from sentence_transformers import SentenceTransformer
except Exception as e:
    raise RuntimeError(
        "chromadb 또는 sentence_transformers 로드 실패. "
        "설치 셀을 먼저 실행했는지 확인하세요."
    ) from e

# =========================================
# 2. rag_documents 로드
# =========================================
conn = get_conn(DB_PATH)

try:
    rag_df = pd.read_sql("SELECT * FROM rag_documents", conn)
finally:
    conn.close()

if rag_df.empty:
    print("rag_documents 테이블이 비어 있습니다.")
    print("먼저 Cell 11을 실행해 risk_event_candidates와 rag_documents를 생성하세요.")
    raise ValueError("rag_documents is empty")

print(f"RAG 문서 로드 완료: {len(rag_df):,}건")

# metadata_json 파싱
def parse_metadata_json_safe(x):
    try:
        return json.loads(x)
    except Exception:
        return {}

rag_df["metadata"] = rag_df["metadata_json"].apply(parse_metadata_json_safe)

# 검색 편의를 위해 주요 metadata 컬럼 확장
rag_df["article_id"] = rag_df["metadata"].apply(lambda m: str(m.get("article_id", "")))
rag_df["title"] = rag_df["metadata"].apply(lambda m: str(m.get("title", "")))
rag_df["outlet"] = rag_df["metadata"].apply(lambda m: str(m.get("outlet", "")))
rag_df["published_date"] = rag_df["metadata"].apply(lambda m: str(m.get("published_date", "")))
rag_df["domains_rule"] = rag_df["metadata"].apply(lambda m: str(m.get("domains_rule", "")))
rag_df["event_type"] = rag_df["metadata"].apply(lambda m: str(m.get("event_type", "")))
rag_df["matched_keywords"] = rag_df["metadata"].apply(lambda m: str(m.get("matched_keywords", "")))
rag_df["risk_level_rule"] = rag_df["metadata"].apply(lambda m: str(m.get("risk_level_rule", "")))
rag_df["url"] = rag_df["metadata"].apply(lambda m: str(m.get("url", "")))

# =========================================
# 3. 사용자 질의 입력
# =========================================
vector_query = input("벡터 검색 질의 (Enter = 기본: PF): ").strip()

if not vector_query:
    vector_query = "PF"

print(f"\n입력 질의: {vector_query}")

# =========================================
# 4. Hybrid Prefilter 함수
# =========================================
def is_pf_query(q: str) -> bool:
    q = str(q or "").upper()
    return bool(re.search(r"(?<![A-Z])PF(?![A-Z])", q)) or "부동산PF" in q or "부동산 PF" in q


def contains_pf_context(text: str) -> bool:
    """
    PF 검색 오탐 방지용.
    SPF, KCP, CP 등 단순 포함 오탐을 줄이고,
    부동산 PF/프로젝트금융/브릿지론 등 맥락을 우선 인정한다.
    """
    text = str(text or "")

    pf_patterns = [
        r"(?<![A-Za-z])PF(?![A-Za-z])",
        r"부동산\s*PF",
        r"프로젝트\s*파이낸싱",
        r"프로젝트파이낸싱",
        r"프로젝트금융",
        r"PF\s*수수료",
        r"PF\s*만기",
        r"PF\s*대출",
        r"PF\s*사업",
        r"PF\s*부실",
        r"PF\s*시장",
        r"PF\s*리스크",
        r"브릿지론",
        # 아래 일반어는 단독으로는 PF 맥락으로 인정하지 않음: 시행사, 차주
        r"만기연장수수료",
        r"페널티수수료",
    ]

    return any(re.search(p, text, flags=re.IGNORECASE) for p in pf_patterns)


def hybrid_prefilter(df: pd.DataFrame, query: str) -> pd.DataFrame:
    """
    질의 유형별 1차 필터링.
    - PF 질의: 부동산 PF 리스크 / PF 맥락 문서만 우선 선별
    - 공사비/자재 질의: 공사비·자재비 상승, 자재 공급망 불안 우선
    - 금리 질의: 금리·통화정책 우선
    - 노동 질의: 노동·임금·인력 리스크 우선
    - 그 외: 전체 문서 사용
    """
    q = str(query or "")
    q_lower = q.lower()

    working_df = df.copy()

    if is_pf_query(q) or "pf" in q_lower or "프로젝트" in q or "브릿지론" in q:
        mask = (
            working_df["event_type"].str.contains("부동산 PF 리스크", na=False)
            | working_df["matched_keywords"].apply(contains_pf_context)
            | working_df["title"].apply(contains_pf_context)
            | working_df["document_text"].apply(contains_pf_context)
        )

        filtered = working_df[mask].copy()

        if not filtered.empty:
            print(f"Hybrid Prefilter: PF 관련 후보 {len(working_df):,}건 → {len(filtered):,}건")
            return filtered

    if any(k in q for k in ["공사비", "자재", "시멘트", "철근", "레미콘", "원자재"]):
        mask = working_df["event_type"].str.contains("공사비|자재", na=False)
        filtered = working_df[mask].copy()

        if not filtered.empty:
            print(f"Hybrid Prefilter: 공사비/자재 후보 {len(working_df):,}건 → {len(filtered):,}건")
            return filtered

    if any(k in q for k in ["금리", "통화", "한은", "FOMC", "국채"]):
        mask = working_df["event_type"].str.contains("금리|통화", na=False)
        filtered = working_df[mask].copy()

        if not filtered.empty:
            print(f"Hybrid Prefilter: 금리/통화 후보 {len(working_df):,}건 → {len(filtered):,}건")
            return filtered

    if any(k in q for k in ["노동", "파업", "임금", "노조", "인력"]):
        mask = working_df["event_type"].str.contains("노동|임금|인력", na=False)
        filtered = working_df[mask].copy()

        if not filtered.empty:
            print(f"Hybrid Prefilter: 노동시장 후보 {len(working_df):,}건 → {len(filtered):,}건")
            return filtered

    print(f"Hybrid Prefilter: 적용 조건 없음 → 전체 {len(working_df):,}건 사용")
    return working_df


# =========================================
# 5. 검색 대상 문서 선정
# =========================================
if USE_HYBRID_PREFILTER:
    search_df = hybrid_prefilter(rag_df, vector_query)
else:
    search_df = rag_df.copy()

if search_df.empty:
    print("검색 대상 문서가 없습니다.")
    raise ValueError("search_df is empty")

# 같은 article_id가 여러 event_type으로 존재할 수 있음.
# 단, PF 질의에서는 부동산 PF 리스크 row를 우선 남긴다.
def event_priority(row, query):
    event_type = str(row.get("event_type", ""))
    matched_keywords = str(row.get("matched_keywords", ""))
    title = str(row.get("title", ""))
    text = str(row.get("document_text", ""))

    if is_pf_query(query):
        if "부동산 PF 리스크" in event_type:
            return 0
        if contains_pf_context(title) or contains_pf_context(matched_keywords) or contains_pf_context(text):
            return 1
        return 9

    # 일반 질의에서는 rule score 컬럼이 없으므로 event_type 있는 문서 우선
    if event_type:
        return 1
    return 9

search_df["event_priority"] = search_df.apply(lambda row: event_priority(row, vector_query), axis=1)

search_df = (
    search_df
    .sort_values(["article_id", "event_priority"])
    .drop_duplicates(subset=["article_id"], keep="first")
    .reset_index(drop=True)
)

print(f"article_id 기준 중복 제거 후 검색 대상: {len(search_df):,}건")

# =========================================
# 6. 임베딩 모델 로드 및 ChromaDB 생성
# =========================================
embed_model = SentenceTransformer(
    "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2"
)

chroma_client = chromadb.PersistentClient(path=VECTOR_DB_RUNTIME_PATH)

collection = chroma_client.get_or_create_collection(
    name=CHROMA_COLLECTION_NAME
)

doc_texts = search_df["document_text"].fillna("").astype(str).tolist()
doc_ids = search_df["doc_id"].fillna("").astype(str).tolist()

# doc_id 중복/공백 보정
fixed_doc_ids = []
seen_doc_ids = set()

for i, doc_id in enumerate(doc_ids):
    doc_id = str(doc_id).strip()

    if not doc_id or doc_id in seen_doc_ids:
        doc_id = f"rag_doc_{i}_{int(time.time())}"

    seen_doc_ids.add(doc_id)
    fixed_doc_ids.append(doc_id)

metadatas = []

for _, row in search_df.iterrows():
    raw_meta = row.get("metadata_json", "{}")

    try:
        meta = json.loads(raw_meta)
    except Exception:
        meta = {}

    # 출력 품질을 위해 보정된 대표 event_type 반영
    meta["event_type"] = row.get("event_type", "")
    meta["matched_keywords"] = row.get("matched_keywords", "")
    meta["article_id"] = row.get("article_id", "")
    meta["title"] = row.get("title", "")
    meta["outlet"] = row.get("outlet", "")
    meta["published_date"] = row.get("published_date", "")
    meta["domains_rule"] = row.get("domains_rule", "")
    meta["risk_level_rule"] = row.get("risk_level_rule", "")
    meta["url"] = row.get("url", "")

    clean_meta = {
        str(k): "" if v is None else str(v)
        for k, v in meta.items()
    }

    metadatas.append(clean_meta)

embeddings = embed_model.encode(
    doc_texts,
    batch_size=32,
    show_progress_bar=True
).tolist()

collection.add(
    ids=fixed_doc_ids,
    documents=doc_texts,
    metadatas=metadatas,
    embeddings=embeddings
)

print(f"ChromaDB 저장 완료: {len(doc_texts):,}건")

# =========================================
# 7. Vector Search 실행
# =========================================
query_embedding = embed_model.encode([vector_query]).tolist()[0]

search_results = collection.query(
    query_embeddings=[query_embedding],
    n_results=min(20, len(doc_texts)),
    include=["documents", "metadatas", "distances"]
)

# =========================================
# 8. 결과 출력
# =========================================
print("\n" + "=" * 60)
print(f"[Hybrid Vector Search 결과: '{vector_query}']")
print("=" * 60)

result_documents = search_results.get("documents", [[]])[0]
result_metadatas = search_results.get("metadatas", [[]])[0]
result_distances = search_results.get("distances", [[]])[0]

display_rank = 1
seen_article_ids = set()

for doc, meta, distance in zip(result_documents, result_metadatas, result_distances):
    article_id = meta.get("article_id", "")

    if article_id and article_id in seen_article_ids:
        continue

    if article_id:
        seen_article_ids.add(article_id)

    # PF 질의에서는 출력 단계에서도 한 번 더 필터
    if is_pf_query(vector_query):
        check_text = " ".join([
            str(meta.get("title", "")),
            str(meta.get("event_type", "")),
            str(meta.get("matched_keywords", "")),
            str(doc),
        ])

        if not contains_pf_context(check_text):
            continue

    # document_text 안에 URL이 이미 들어 있으므로 출력용 판단 근거에서는 URL 라인 제거
    clean_doc_lines = [
        line for line in str(doc).splitlines()
        if not line.strip().startswith("[URL]")
    ]
    clean_doc = "\n".join(clean_doc_lines)

    print(f"\n--- 결과 {display_rank} ---")
    print(f"[기사 제목] {meta.get('title', '')}")
    print(f"[언론사] {meta.get('outlet', '')}")
    print(f"[일자] {meta.get('published_date', '')}")
    print(f"[도메인] {meta.get('domains_rule', '')}")
    print(f"[위험 이벤트] {meta.get('event_type', '')}")
    print(f"[매칭 키워드] {meta.get('matched_keywords', '')}")
    print(f"[거리] {float(distance):.4f}")
    print(f"[판단 근거]\n{clean_doc[:500]}")
    print(f"[URL] {meta.get('url', '')}")

    display_rank += 1

    if display_rank > TOP_K_DISPLAY:
        break

if display_rank == 1:
    print("검색 결과가 없습니다.")
    print("질의를 조금 더 구체화해보세요. 예: '부동산 PF 수수료', '브릿지론', 'PF 만기연장'")

print("\nCell 16 실행 완료")

# %% [markdown]
"""
## Conclusion

본 PoC는 BigKinds 1일치 종합일간지 뉴스 데이터를 대상으로 EWS 관련 4개 도메인별 위험 후보 언론보도 데이터셋을 생성할 수 있음을 검증합니다.

운영 단계에서는 원천 수집 자동화, 저작권/이용권한 검토, 도메인 사전 고도화, Airflow 기반 스케줄링, LLM 판단 평가 체계가 추가로 필요합니다.

"""
