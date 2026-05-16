# 근로계약서 생성기

경동TPR 근로계약서 양식을 웹 폼으로 작성하면 DOCX 파일로 다운로드해주는 Flask 웹 앱입니다.

## 실행 방법

```bash
# 1. 가상환경 생성 및 활성화
python3 -m venv venv
source venv/bin/activate       # Windows: venv\Scripts\activate

# 2. 패키지 설치
pip install -r requirements.txt

# 3. 서버 실행
python app.py
```

브라우저에서 `http://localhost:5000` 접속 후 양식을 작성하고 버튼을 누르면 계약서 DOCX 파일이 자동 다운로드됩니다.

## 프로젝트 구조

```
├── app.py                  # Flask 라우터
├── generator.py            # DOCX 생성 로직
├── templates/
│   └── index.html          # 웹 폼 UI
├── static/
│   └── style.css           # 스타일
├── 근로계약서_양식.docx      # 원본 템플릿
└── requirements.txt
```
