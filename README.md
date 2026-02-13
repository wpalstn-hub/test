# 발주 엑셀 자동 병합 스크립트

`merge_orders_to_master.ps1`는 Windows Excel COM 자동화를 사용해 아래 작업을 수행합니다.

- 발주 폴더(`*.xlsx`)의 파일들을 순서대로 읽음
- 각 파일의 헤더 행(상위 10행에서 가장 일치율 높은 행)을 자동 탐지
- `전체 제품 관리 리스트.xlsb`의 1행 헤더와 일치하는 열만 선택
- 최신 데이터를 마스터 시트의 마지막 행 아래에 추가

## 사용 방법

PowerShell에서 아래처럼 실행하세요.

```powershell
powershell -ExecutionPolicy Bypass -File .\merge_orders_to_master.ps1 `
  -OrderFolder "C:\Users\wpals\OneDrive\바탕 화면\CODEX TEST\발주" `
  -MasterWorkbook "C:\Users\wpals\OneDrive\바탕 화면\CODEX TEST\전체 제품 관리 리스트.xlsb"
```

특정 시트만 대상으로 하려면:

```powershell
powershell -ExecutionPolicy Bypass -File .\merge_orders_to_master.ps1 `
  -OrderFolder "...\발주" `
  -MasterWorkbook "...\전체 제품 관리 리스트.xlsb" `
  -MasterSheetName "Sheet1"
```

## 주의 사항

- Windows + Microsoft Excel 설치 환경에서 실행해야 합니다.
- 병합 전 원본 파일 백업을 권장합니다.
- 헤더는 기본적으로 마스터 파일의 **1행** 기준입니다.
