# 업무용승용차 운행기록부 자동 업데이트 앱

## 수정 사항
- .xls 손상/비표준 파일 대응
- 병합 셀(MergedCell) 쓰기 오류 대응
- 병합 영역 내부 셀에 값을 쓸 때 자동으로 왼쪽 위 셀에 입력

## 실행 방법

```bash
pip install -r requirements.txt
streamlit run app.py
```
