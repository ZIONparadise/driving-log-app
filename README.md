# 업무용승용차 운행기록부 자동 업데이트 앱

## 이번 버전 수정

- 주행거리(S열): `=O행-K행`
- 일반 업무용(AA열): `=S행-W행`
- Excel 파일을 열 때 수식이 자동 재계산되도록 설정
- 저장 파일명 자동 생성: `업무용운행기록부_30th_홍성학_5월.xlsx`

## 참고
openpyxl은 수식 결과를 직접 계산하지 않습니다.
그래서 Excel에서 파일을 열 때 자동 계산되도록 fullCalcOnLoad / forceFullCalc / calcMode=auto 설정을 적용했습니다.
