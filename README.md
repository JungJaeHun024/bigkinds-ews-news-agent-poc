# BigKinds 기반 EWS 언론보도 AI Agent PoC



본 프로젝트는 건설산업 조기경보시스템(EWS)에서 활용 가능한 비정형 언론보도 데이터를 구축하기 위해, BigKinds 뉴스 데이터를 기반으로 도메인 분류·위험 이벤트 후보 추출·LLM 최종 판단을 수행하는 AI Agent PoC입니다.



본 PoC의 목적은 완성형 뉴스 검색엔진 개발이 아니라, BigKinds Excel 다운로드 데이터를 활용해 건설경기·금융시장·건설자재·노동시장 4개 도메인별 언론보도 후보 데이터셋을 구조화할 수 있는지 검증하는 것입니다.



## 1. 문제 정의



건설산업 조기경보시스템은 정형 지표뿐 아니라 언론보도와 같은 비정형 데이터에서 나타나는 위험 신호를 함께 모니터링할 필요가 있습니다.



그러나 언론보도 데이터는 기사 제목, 본문, 키워드, 출처 URL 등 비정형 요소가 혼재되어 있어 단순 수집만으로는 EWS에 바로 활용하기 어렵습니다.



본 PoC는 다음 질문을 검증합니다.



\- BigKinds 뉴스 Excel 데이터를 EWS용 원천 데이터로 활용할 수 있는가?

\- AI Agent가 기사를 4개 도메인으로 분류할 수 있는가?

\- 위험 이벤트 후보를 구조화된 DB 테이블로 저장할 수 있는가?

\- LLM을 전수 판단기가 아니라 상위 후보 검토 보조 모듈로 활용할 수 있는가?



## 2. 데이터 수집 범위



\- 원천: BigKinds 뉴스검색 결과 Excel

\- 수집 방식: 웹사이트 검색 후 Excel 다운로드 기반 반자동 수집

\- 수집 대상: 1일치 종합일간지 뉴스 기사

\- PoC 데이터 규모: 3,096건

- 대상 도메인:

&#x20; - 건설경기

&#x20; - 금융시장

&#x20; - 건설자재

&#x20; - 노동시장



원천 BigKinds Excel 파일은 저작권 및 이용 범위 검토가 필요한 데이터이므로 저장소에는 포함하지 않습니다.



## 3. Agent Pipeline



본 PoC는 다음 흐름으로 구성됩니다.



1\. BigKinds Excel Ingestion

2\. 기사 컬럼 표준화 및 DB 적재

3\. 도메인 키워드 사전 기반 1차 분류

4\. 위험 이벤트 후보 탐지

5\. RAG 문서 테이블 생성

6\. 사용자 자연어 질의 기반 후보 기사 조회

7\. NVIDIA LLM 기반 최종 판단

8\. Excel 결과 리포트 생성



## 4. PoC 결과



| 항목 | 결과 |

|---|---:|

| 원천 기사 수 | 3,096건 |

| 도메인 태깅 결과 | 862건 |

| Rule 기반 위험 후보 | 534건 |

| 이벤트 요약 | 10개 |

| RAG 문서 | 534건 |

| LLM 최종 판단 샘플 | 수행 완료 |



본 결과를 통해 BigKinds 기반 언론보도 데이터를 EWS 도메인별 위험 후보 데이터셋으로 구조화할 수 있음을 확인했습니다.



## 5. 차별화 포인트



### 5.1 산업 도메인 특화



본 PoC는 범용 챗봇이나 단순 뉴스 요약이 아니라, 건설산업 EWS 관점에서 필요한 도메인 구조를 먼저 정의했습니다.



건설경기·금융시장·건설자재·노동시장 키워드 사전을 직접 설계하고, 부동산 PF와 같은 산업 특화 위험 이벤트를 별도 분류했습니다.



특히 PF 검색 과정에서 SPF, KCP, CP 등 유사 토큰으로 인한 오탐 가능성을 확인하고, PF 맥락을 구분하는 예외 처리 로직을 반영했습니다.



### 5.2 트러블슈팅 기반 개선



초기 LLM 호출 과정에서 API timeout 문제가 발생했습니다.



이를 해결하기 위해 batch size를 축소하고, 빠른 모델 옵션과 streaming 응답 방식을 추가했습니다.



이 과정은 설정값과 코드 주석으로 남겨, PoC 과정에서 발생한 성능·안정성 문제와 해결 근거를 추적할 수 있도록 했습니다.



### 5.3 명확한 PoC 목적



본 프로젝트는 운영용 뉴스 검색엔진을 개발하는 것이 아니라, 언론보도 데이터의 수집 가능성과 EWS 데이터셋 구조화 가능성을 판단하기 위한 PoC입니다.



따라서 LLM을 전체 기사 전수 판단에 사용하지 않고, Rule 기반 후보 선별 이후 상위 후보에 대한 최종 판단 보조 역할로 제한했습니다.



## 6. 한계 및 운영 확장 방향



\- 원천 수집은 BigKinds 웹사이트 Excel 다운로드 기반 반자동 구조입니다.

\- 기사 원문 저장 및 재활용 범위는 저작권 검토가 필요합니다.

\- 키워드 기반 후보 탐지는 오탐 가능성이 있으므로 운영 단계에서는 품질 검증 체계가 필요합니다.

\- 운영 단계에서는 Airflow 기반 정기 적재, 중복 제거, 모니터링, 실패 재처리 구조가 필요합니다.

\- Vector Search는 부가 검증 기능이며, 운영 검색엔진 수준으로 사용하려면 별도 고도화가 필요합니다.



## 7. Repository Structure



```text

bigkinds-ews-news-agent-poc/

├─ README.md

├─ requirements.txt

├─ notebooks/

│  └─ ews\_bigkinds\_agent\_poc.ipynb

├─ src/

│  └─ ews\_bigkinds\_agent\_poc.py

├─ docs/

├─ outputs/

└─ data/

&#x20;  ├─ raw/

&#x20;  └─ sample/
```

## 8. 실행 방법

### 8.1 실행 환경 준비

본 프로젝트는 Google Colab 또는 로컬 Python 환경에서 실행할 수 있습니다.

권장 실행 환경은 다음과 같습니다.

- Python 3.10 이상
- Jupyter Notebook 또는 Google Colab
- NVIDIA API Key
- BigKinds 뉴스검색 결과 Excel 파일

### 8.2 의존성 설치

로컬 환경에서 실행하는 경우 다음 명령어로 필요한 패키지를 설치합니다.

    pip install -r requirements.txt

Google Colab에서 실행하는 경우 노트북의 설치 셀을 먼저 실행합니다.

### 8.3 입력 데이터 준비

BigKinds 웹사이트에서 뉴스검색 결과 Excel 파일을 다운로드합니다.

본 PoC에서는 1일치 종합일간지 뉴스 데이터를 사용했습니다.

예시 파일명은 다음과 같습니다.

    NewsResult_YYYYMMDD-YYYYMMDD.xlsx

원천 BigKinds Excel 파일은 저작권 및 이용 범위 검토가 필요한 데이터이므로 본 저장소에는 포함하지 않습니다.

### 8.4 노트북 실행

다음 노트북을 실행합니다.

    notebooks/ews_bigkinds_agent_poc.ipynb

노트북의 설정 셀에서 원천 Excel 파일 경로를 지정합니다.

예시:

    FILE_PATH = '/content/NewsResult_20260517-20260518.xlsx'

또는 환경변수를 사용할 수 있습니다.

    BIGKINDS_FILE_PATH

### 8.5 API Key 입력

NVIDIA API Key는 코드에 직접 저장하지 않습니다.

노트북 실행 중 다음 방식으로 입력합니다.

    NVIDIA_API_KEY 입력:

API Key는 .env, notebook output, GitHub repository에 포함하지 않습니다.

### 8.6 주요 실행 단계

노트북은 다음 순서로 실행됩니다.

1. 환경 설정
2. BigKinds Excel 파일 로드
3. 기사 컬럼 표준화
4. SQLite DB 적재
5. 도메인 분류
6. 위험 이벤트 후보 생성
7. RAG 문서 테이블 생성
8. 자연어 질의 기반 후보 기사 조회
9. LLM 최종 판단
10. Excel 결과 리포트 생성
11. 검증 리포트 출력
12. Optional Hybrid Vector Search 실행

### 8.7 주요 출력 파일

실행 후 생성되는 주요 파일은 다음과 같습니다.

| 파일 | 설명 |
|---|---|
| ews_bigkinds_agent.db | SQLite 기반 PoC DB |
| ews_bigkinds_agent_report.xlsx | 기본 모니터링 결과 리포트 |
| ews_interactive_report.xlsx | 자연어 질의 기반 결과 리포트 |
| rag_documents | RAG 검색용 문서 테이블 |
| llm_final_judgements | LLM 최종 판단 결과 테이블 |

단, DB 파일과 결과 Excel 파일은 기사 본문 또는 저작권 검토 대상 텍스트를 포함할 수 있으므로 공개 저장소에는 기본적으로 포함하지 않습니다.

## 9. 저장소 구성

본 저장소의 기본 구조는 다음과 같습니다.

    bigkinds-ews-news-agent-poc/
    ├─ README.md
    ├─ requirements.txt
    ├─ notebooks/
    │  └─ ews_bigkinds_agent_poc.ipynb
    ├─ src/
    │  └─ ews_bigkinds_agent_poc.py
    ├─ docs/
    │  ├─ architecture.md
    │  ├─ troubleshooting.md
    │  └─ data_dictionary.md
    ├─ outputs/
    │  └─ result_summary.md
    └─ data/
       ├─ README.md
       ├─ raw/
       └─ sample/

## 10. 포함하지 않는 파일

다음 파일은 공개 저장소에 포함하지 않습니다.

| 제외 대상 | 제외 사유 |
|---|---|
| BigKinds 원천 Excel | 뉴스 기사 저작권 및 이용권한 검토 필요 |
| SQLite DB 파일 | 기사 본문 및 중간 처리 결과 포함 가능 |
| ChromaDB 저장소 | 실행 환경별 생성 파일 |
| API Key | 보안 정보 |
| 대용량 결과 Excel | 기사 본문 포함 가능 |

## 11. 프로젝트 결론

본 PoC를 통해 BigKinds Excel 다운로드 데이터를 기반으로 건설산업 EWS의 4개 도메인별 언론보도 후보 데이터셋을 구축할 수 있음을 확인했습니다.

핵심 결론은 다음과 같습니다.

- BigKinds 뉴스 Excel은 EWS 비정형 언론보도 원천 데이터로 활용 가능합니다.
- 다운로드 이후 전처리, 도메인 분류, 위험 이벤트 후보 생성, DB 적재는 자동화할 수 있습니다.
- LLM은 전체 기사 전수 판단보다 상위 후보 기사에 대한 최종 검토 보조 역할이 적절합니다.
- 운영 단계에서는 수집 자동화, 저작권 검토, 품질 검증, Airflow 기반 정기 적재 구조가 추가로 필요합니다.

## 12. Portfolio Point

본 프로젝트는 단순 뉴스 요약 또는 챗봇 구현이 아니라, 특정 산업 도메인에서 비정형 데이터를 업무 목적에 맞게 구조화하는 과정을 검증한 PoC입니다.

특히 다음 세 가지 점에서 차별성을 가집니다.

1. 건설경기, 금융시장, 건설자재, 노동시장 도메인 사전을 직접 설계했습니다.
2. PF 오탐, LLM timeout, ChromaDB 저장 오류 등 실제 구현 과정의 문제를 해결했습니다.
3. 개발 자체가 아니라 수집 가능성 판단과 EWS 데이터셋 구축 가능성 검증이라는 명확한 목적에 맞춰 구현 범위를 제한했습니다.
