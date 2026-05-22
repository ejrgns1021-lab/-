# 🌤 제주 기상데이터 자동 수집기

기상청 조건별통계에서 **제주·서귀포·성산·고산** 4개 지점의 기상 데이터를  
매일 자동으로 수집하여 엑셀 파일로 저장하는 GitHub Actions 프로젝트입니다.

---

## 📁 파일 구조

```
├── .github/
│   └── workflows/
│       └── daily_weather.yml   ← GitHub Actions 워크플로우
├── scripts/
│   └── fetch_weather.py        ← 데이터 수집 + 엑셀 생성 스크립트
├── data/
│   └── jeju_weather_202506.xlsx  ← 자동 생성되는 월별 엑셀
├── requirements.txt
└── README.md
```

---

## 🚀 설치 방법

### 1단계: 저장소 생성
GitHub에서 새 저장소(public 또는 private)를 만든 뒤  
이 폴더의 파일들을 그대로 업로드합니다.

```bash
git init
git add .
git commit -m "초기 설정"
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO.git
git push -u origin main
```

### 2단계: Actions 권한 확인
저장소 → **Settings → Actions → General**  
→ "Workflow permissions" 에서 **Read and write permissions** 선택 후 저장

### 3단계: 완료 🎉
이후 매일 **오전 9시(KST)** 에 자동 실행됩니다.

---

## ⚙️ 실행 일정

| 구분 | 내용 |
|------|------|
| 자동 실행 | 매일 오전 9시 KST (UTC 00:00) |
| 수동 실행 | Actions 탭 → `제주 기상데이터 자동 수집` → `Run workflow` |
| 수집 기간 | 당월 1일 ~ 전날 |

---

## 📊 수집 데이터 항목

| 요소 | 세부 항목 |
|------|-----------|
| 기온 | 평균·최고·최저 기온(℃) |
| 강수 | 강수량(mm), 1시간최다강수량, 강우일수 |
| 바람 | 평균·최대·최대순간 풍속(m/s), 시각 |
| 습도 | 평균·최저 습도(%rh) |
| 일조일사 | 일조합(hr), 일조율(%), 일사합(MJ/m²) |

엑셀은 **지점별 시트** (제주 / 서귀포 / 성산 / 고산) 로 구성됩니다.

---

## 🗂 출력 파일

`data/jeju_weather_YYYYMM.xlsx` — 월별 1개 파일로 누적 갱신됩니다.  
예) `data/jeju_weather_202506.xlsx`

---

## ❓ 문제 해결

| 증상 | 원인 및 해결 |
|------|-------------|
| Actions가 실행 안 됨 | Settings → Actions → General → 권한 확인 |
| 데이터가 비어 있음 | 기상청 서버 일시 장애 — 다음날 재시도됨 |
| push 권한 오류 | Workflow permissions을 Read/Write로 변경 |
