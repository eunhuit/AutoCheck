# Google Sheets 근태 관리 시스템 📅⏱️

![Python Version](https://img.shields.io/badge/Python-3.12%2B-blue?logo=python)
![License](https://img.shields.io/badge/License-Apache%202.0-green)

> 구글 스프레드시트 기반의 자동화 근태 기록 관리 프로그램

## 🌟 주요 기능
- **실시간 시트 연동**
  - 월별/주차별 자동 시트 생성
  - 출퇴근 시간 자동 기록
  - 외출 시간대 및 사유 기록
- **스마트 시간 계산**
```Python
def gsn(dt): # "월/주차" 계산
def gti(dt): # 테이블 인덱스
산 def ft(dt): # 포맷팅된 시간 문자
```
- **자동화 시트 관리**
  - 새 달 시작 시 시트 자동 생성
  - 부서명 자동 입력(IT네트워크시스템)
- **GUI 인터페이스**
  - 출근/퇴근/외출 버튼 통합 관리
  - 팝업 알림 시스템

## 🛠️ 사용 방법
```bash
pip install gspread oauth2client tkinter
```
## Google API 설정
1. Google Cloud Console에서 [서비스 계정](https://console.cloud.google.com/apis/credentials?inv=1&invt=Abs-SQ&project=flawless-star-346013) 생성
2. Google Sheets [API](https://console.cloud.google.com/marketplace/product/google/sheets.googleapis.com?q=search&referrer=search&inv=1&invt=Abs-SQ&project=flawless-star-346013) 활성화
3. 서비스 계정 키(json 파일) 다운로드
4. 스프레드시트 공유 설정(공개된 시트는 않해도 괜찮습니다!!)

## ⚙️ 설정 가이드
```Python
SERVICE_ACCOUNT_FILE = "./google.json" # 서비스 계정 키 경로
SPREADSHEET_URL = "스프레드시트 URL" # 실제 사용 시 변경 필요
BASE_ROW = 25 # 데이터 입력 행 설정
TABLE_WIDTH = 6 # 주간 테이블 간격
```

## 🖥️ 사용 방법
1. **출근 처리**
   - [출근] 버튼 클릭 → 현재 시간 기록
   - 부서명 자동 입력(IT네트워크시스템)

2. **퇴근 처리**
   - [퇴근] 버튼 클릭 → 퇴근 시간 업데이트

3. **외출 관리**
```Python
외출 시작 → 복귀 시 사유 입력 팝업
def outside(): # 외출 시간 기록
def rfo(): # 복귀 시간 및 사유 처리
```

## 📊 데이터 구조
| 컬럼       | 내용                | 계산 방식                |
|------------|---------------------|-------------------------|
| B25        | 부서명              | 고정값 입력             |
| C25        | 1주차 출근         | BASE_COL_CHECKIN + (주차-1)*6 |
| D25        | 1주차 퇴근         | BASE_COL_CHECKOUT + (주차-1)*6 |
| E25        | 1주차 외출         | BASE_COL_GOOUT + (주차-1)*6 |

## 📦 패키지
| 패키지         | 용도                   |
|----------------|------------------------|
| gspread        | Google Sheets API 연동 |
| oauth2client   | Google 인증           |
| tkinter        | GUI 인터페이스         |
| datetime       | 시간 계산              |

## 📜 라이센스
Apache License 2.0  

## ⚠️ 주의사항
1. **Google API 제한**
   - 1분당 60회 요청 제한
   - 동시 사용자 수 제한(100명)

2. **시간 계산 규칙**
   - 주차 계산: `(day - 1) // 7 + 1`
   - 테스트용 1일 오프셋 포함(`timedelta(days=1)`)

3. **보안 요구사항**
   - google.json 파일 보안 유지
   - 스프레드시트 공유 범위 제한

> 🚨 중요: 실제 사용 전 BASE_ROW 및 DEPARTMENT_NAME 수정 필요

