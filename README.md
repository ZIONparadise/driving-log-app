# 업무용승용차 운행기록부 자동 업데이트 앱

## 기능
- 제출용 운행기록부 양식 업로드
- 다운로드 받은 운행내역 파일(.xls/.xlsx) 업로드
- 과세기간, 사용일자, 요일, 부서, 성명, 주행 전/후 거리, 주행거리, 업무용 사용거리, 비고 자동 입력
- 업데이트된 엑셀 파일 다운로드

## 실행 방법

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud 배포
GitHub 저장소에 아래 파일을 업로드하세요.

- app.py
- requirements.txt

그 다음 Streamlit Cloud에서 `app.py`를 선택해 배포하면 됩니다.
