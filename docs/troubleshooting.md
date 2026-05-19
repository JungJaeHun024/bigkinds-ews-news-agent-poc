# Troubleshooting Notes



## 1. LLM API Timeout



### Problem



NVIDIA LLM API 호출 시 후보 기사 수가 많거나 prompt가 길어지면 timeout이 발생했습니다.



### Cause



주요 원인은 다음과 같습니다.



\- 70B 모델 사용

\- 긴 기사 본문 입력

\- 큰 batch size

\- 긴 응답 생성

\- Colab 환경의 네트워크 지연



\### Solution



다음 방식으로 해결했습니다.



\- 빠른 모델 옵션 추가

\- batch size 축소

\- 기사 본문 입력 길이 제한

\- streaming 응답 방식 적용

\- Rule 기반 후보 선별 후 LLM 최종 판단 수행



## 2. PF Keyword False Positive



### Problem



PF 검색 과정에서 SPF, KCP, CP 등 유사 토큰으로 인해 비관련 기사가 후보에 포함될 수 있었습니다.



### Cause



단순 문자열 포함 방식으로 PF를 탐지하면 SPF, KCP, CP와 같은 단어 내부의 문자열도 함께 매칭됩니다.



### Solution



다음 로직을 추가했습니다.



\- PF 단독 토큰 정규식 적용

\- 부동산 PF, 프로젝트금융, 브릿지론 등 맥락 키워드 추가

\- metadata 기반 prefilter 적용

\- Vector Search는 Optional 기능으로 분리



## 3. ChromaDB Readonly Error



### Problem



Colab 환경에서 ChromaDB를 반복 생성할 때 readonly database 오류가 발생했습니다.



### Cause



기존 ChromaDB path를 반복 재사용하면서 SQLite lock 또는 권한 문제가 발생했습니다.



### Solution



다음 방식으로 해결했습니다.



\- 매 실행마다 runtime path를 새로 생성

\- 기존 persistent DB 재사용을 피함

\- ChromaDB 기능은 핵심 파이프라인이 아니라 Optional 검증 셀로 분리



## 4. PoC Scope Control



### Problem



Vector Search 품질을 계속 고도화하려 하면 PoC 범위를 넘어 운영형 검색엔진 개발로 확장될 수 있었습니다.



### Decision



본 프로젝트의 목적은 완성형 검색엔진 개발이 아니라 BigKinds 언론보도 데이터의 EWS 데이터셋 구조화 가능성 검증입니다.



따라서 Vector Search는 부가 검증으로 제한하고, 핵심 산출물은 Rule 기반 도메인 분류와 위험 이벤트 후보 데이터셋으로 정의했습니다.

