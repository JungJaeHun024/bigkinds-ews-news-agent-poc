# Architecture



## 1. Overview



본 프로젝트는 BigKinds Excel 뉴스 데이터를 건설산업 EWS용 언론보도 후보 데이터셋으로 구조화하는 AI Agent PoC입니다.



Agent는 원천 기사 데이터를 직접 최종 판단하지 않고, Rule 기반 후보 선별 이후 필요한 일부 후보에 대해서만 LLM 최종 판단을 수행합니다.



## 2. Agent Workflow



처리 흐름은 다음과 같습니다.



1\. BigKinds Excel 입력

2\. Ingestion Agent

3\. Raw Article DB 저장

4\. Domain Classification

5\. Risk Event Candidate Detection

6\. RAG Document Builder

7\. Query Planner

8\. Tool Executor

9\. LLM Final Judgement

10\. Excel Report 생성



## 3. Table Design



| Table | Description |

|---|---|

| raw\_articles | BigKinds 원천 기사 저장 |

| article\_domain\_tags | 4개 도메인 분류 결과 저장 |

| risk\_event\_candidates | 위험 이벤트 후보 저장 |

| event\_summary | 이벤트 유형별 요약 결과 저장 |

| rag\_documents | RAG 검색용 문서 저장 |

| llm\_final\_judgements | LLM 최종 판단 결과 저장 |

| query\_logs | 사용자 질의 로그 저장 |

| agent\_outputs | Agent 실행 결과 저장 |



## 4. Design Principle



본 PoC는 LLM을 전체 기사 전수 판단에 사용하지 않습니다.



먼저 Rule 기반 도메인 사전과 위험 이벤트 사전으로 후보 기사를 선별하고, 이후 LLM은 상위 후보 기사에 대한 최종 판단 보조 역할을 수행합니다.



이 구조는 다음 목적을 가집니다.



\- LLM 호출 비용 감소

\- 응답 속도 개선

\- Rule 기반 데이터 구조화 가능성 확보

\- 최종 판단 책임 범위 제한

\- EWS 데이터셋 구축 목적에 맞는 PoC 범위 유지



## 5. Scope



본 프로젝트는 운영 시스템이 아니라 PoC입니다.



따라서 다음은 포함하지 않습니다.



\- BigKinds 자동 크롤링

\- 실시간 수집 자동화

\- 운영용 검색엔진

\- 전체 기사 LLM 전수 판단

\- 배포형 웹 서비스



