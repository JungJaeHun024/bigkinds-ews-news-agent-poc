\# Data Dictionary



\## 1. raw\_articles



BigKinds Excel에서 적재한 원천 기사 테이블입니다.



| Column | Description |

|---|---|

| article\_id | 내부 기사 ID |

| bigkinds\_news\_id | BigKinds 기사 ID |

| published\_date | 기사 발행일 |

| outlet | 언론사 |

| title | 기사 제목 |

| content | 기사 본문 또는 요약 |

| keywords | BigKinds 키워드 |

| features | 특성추출어 |

| url | 원문 URL |

| source\_file | 원천 파일명 |

| loaded\_at | 적재 시각 |



\## 2. article\_domain\_tags



기사별 4개 도메인 분류 결과입니다.



| Column | Description |

|---|---|

| article\_id | 기사 ID |

| domain | Rule 기반 도메인 |

| matched\_keywords | 매칭된 도메인 키워드 |

| domain\_score | 도메인 매칭 점수 |

| created\_at | 생성 시각 |



\## 3. risk\_event\_candidates



위험 이벤트 후보 기사 테이블입니다.



| Column | Description |

|---|---|

| event\_result\_id | 이벤트 결과 ID |

| article\_id | 기사 ID |

| published\_date | 기사 발행일 |

| outlet | 언론사 |

| title | 기사 제목 |

| url | 기사 URL |

| domains\_rule | Rule 기반 도메인 |

| event\_type | 위험 이벤트 유형 |

| matched\_keywords | 매칭 키워드 |

| evidence\_text | 판단 근거 텍스트 |

| risk\_score\_rule | Rule 기반 위험 점수 |

| risk\_level\_rule | Rule 기반 위험 수준 |

| detected\_by | 탐지 방식 |

| created\_at | 생성 시각 |



\## 4. event\_summary



위험 이벤트 유형별 요약 테이블입니다.



| Column | Description |

|---|---|

| published\_date | 기준 일자 |

| event\_type | 위험 이벤트 유형 |

| article\_count | 기사 수 |

| outlet\_count | 언론사 수 |

| max\_rule\_score | 최대 Rule 점수 |

| risk\_level\_summary | 요약 위험 수준 |



\## 5. rag\_documents



RAG 검색용 문서 테이블입니다.



| Column | Description |

|---|---|

| doc\_id | RAG 문서 ID |

| article\_id | 기사 ID |

| document\_text | 검색용 문서 텍스트 |

| metadata\_json | 문서 메타데이터 |

| created\_at | 생성 시각 |



\## 6. llm\_final\_judgements



LLM 최종 판단 결과 테이블입니다.



| Column | Description |

|---|---|

| article\_id | 기사 ID |

| is\_risk\_relevant | 위험 관련성 판단 |

| related\_domains | 관련 도메인 |

| risk\_event\_type | LLM 판단 위험 이벤트 |

| risk\_level | LLM 판단 위험 수준 |

| judgement\_reason | 판단 사유 |

| followup\_required | 후속 검토 필요 여부 |

| judgement\_status | 판단 상태 |

| created\_at | 생성 시각 |

