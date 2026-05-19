# Data Directory



본 저장소는 BigKinds 원천 Excel 파일을 포함하지 않습니다.



## 1. 원천 데이터 제외 사유



BigKinds 뉴스 데이터는 기사 본문, 언론사 메타데이터, URL 등을 포함할 수 있습니다.



따라서 저작권 및 이용권한 검토가 필요한 데이터이므로 공개 저장소에는 포함하지 않습니다.



## 2. 예상 입력 파일



실행 시 사용자는 BigKinds 웹사이트에서 뉴스검색 결과 Excel 파일을 직접 다운로드해야 합니다.



예시 파일명은 다음과 같습니다.



&#x20;   NewsResult\_YYYYMMDD-YYYYMMDD.xlsx



## 3. 사용 방법



노트북의 설정 셀에서 원천 Excel 파일 경로를 지정합니다.



예시:



&#x20;   FILE\_PATH = '/content/NewsResult\_20260517-20260518.xlsx'



또는 환경변수 BIGKINDS\_FILE\_PATH를 사용할 수 있습니다.



## 4. 보관 정책



| 구분 | 저장소 포함 여부 |

|---|---|

| 원천 Excel | 포함하지 않음 |

| SQLite DB | 포함하지 않음 |

| ChromaDB 파일 | 포함하지 않음 |

| 샘플 스키마 | 포함 가능 |

| 결과 요약 문서 | 포함 가능 |

