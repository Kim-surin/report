# PPT Report Data Auto-Filler

## Project Overview

데이터를 조회하여 PPT 보고서 템플릿의 플레이스홀더/표/차트에 데이터를 자동으로 채워 넣는 툴입니다.
PPT 변환은 필요 없으며, 단순 PPT에 데이터만 채워넣으세요.

데이터: report\데이터
보고서: report\보고서

예시. 데이터 폴더의 202601 텍스트 파일은 2026년 1월 보고서에 해당하는 데이터입니다.
보고서 폴더의 202601 파일에 데이터를 채워넣으세요.

대상은 데이터 폴더에 존재하는 전체 텍스트 파일입니다.

## Code Conventions & Rules

- **Template Safety:** 기존 PPTX 파일의 원본 서식(폰트, 크기, 색상, 정렬)이 깨지지 않도록, 텍스트 입력 시 기존 Run의 스타일을 유지하세요.

## Data Mapping Strategy (중요)

- **Table Handling:** PPT 내 표(Table)에 데이터를 채울 때는 셀 단위로 기존 서식을 유지하면서 `text`를 교체하세요.
- **Error Handling:** 데이터 매핑 실패나 PPT shape를 찾지 못한 경우 프로세스가 멈추지 않고 로그(logging)를 남기도록 작성하세요.

## 예시

<1. 원산지 판정 및 비역내산 판정 상세>
PRODUCT_COUNT: 전체 품목 수-판정 건수
COO_Y_COUNT: 역내산 판정 건수-판정 건수
COO_N_COUNT: 비역내산 판정 건수-판정 건수
COO_ERR_COUNT: 판정 실패 건수-판정 건수
COO_YT_RATIO: 전체 품목 수-비율
COO_Y_RATIO: 역내산 비율-비율
COO_N_RATIO: 비역내산 비율-비율
COO_ERROR_RATIO: 판정 실패 건수-비율
CTC_NO_COUNT: 세번변경기준 미충족-판정 건수
RVC_NO_COUNT: 부가가치기준 미충족-판정 건수
COMBINE_NO_COUNT: 조합기준 미충족-판정 건수
COO_N_COUNT2: 비역내산 합계-판정 건수
CTC_NO_RATIO: 세번변경기준 미충족-비율
RVC_NO_RATIO: 부가가치기준 미충족-판정 건수
COMBINE_NO_RATIO: 조합기준 미충족-판정 건수
COO_NT_RATIO: 비역내산 합계-판정 건수

<2. 원산지 판정 에러 현황>
HS코드 누락: HS코드 누락
BOM 누락: BOM 누락
소요량 또는 금액이 0 인 것이 존재합니다.: 원재료 단가
그외 항목: 기타

<3. 협정별 원산지 판정 현황>
FTA_NAME: FTA 협정
PRODUCT_COUNT: 판정대상품목(개)
COO_Y_COUNT: (1)역내산 품목수(개)
COO_N_COUNT: (2)비역내산 품목수(개)
COO_ERR_COUNT: (3)판정에러 품목수(개)
COO_Y_RATIO: 역내산 비율(%)

<1. 협력사별 원산지확인서 수취 현황>
VENDOR_NAME: 협력사명
ITEM_COUNT: 구매 자재수(개)
COO_T_COUNT: 수취 자재수(개)
COO_Y_COUNT: 역내 수취 자재수(개)
COO_N_COUNT: 비역내 수취 자재수(개)
NOT_COUNT: 미수취 자재수(개)
COO_RATIO: 수취율(%)
PO_AMOUNT: 구매금액(원)

<2. 협력사별 원산지확인서 수취/미수취 자재 내역>
VENDOR_NAME: 협력사명
ITEM_CODE: 자재코드
HS_CODE: HS CODE
UNIT_PRICE: 구매단가(원)
PO_AMOUNT: 구매금액(원)
COO_N: 수취여부
