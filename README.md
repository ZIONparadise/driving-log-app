# 업무용승용차 운행기록부 자동 업데이트 앱

## 수정 사항
- 일부 .xls 파일에서 발생하는 `CompDocError: Workbook corruption` 오류 대응
- `ignore_workbook_corruption=True` 적용
- .xls가 실제 HTML table 형식으로 저장된 경우에도 읽기 시도

## 실행 방법

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud 배포
GitHub 저장소에 아래 파일을 업로드하세요.

- app.py
- requirements.txt
- README.md
