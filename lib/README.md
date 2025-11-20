# SheetJS Library - Offline Storage

## 라이브러리 설치 방법 (How to Install the Library)

이 디렉토리에는 Excel 파일 처리를 위한 SheetJS (xlsx) 라이브러리가 저장되어야 합니다.

### Option 1: 온라인 환경에서 다운로드 (Download from Online Environment)

외부 인터넷이 가능한 환경에서 다음 중 하나의 방법으로 다운로드:

1. **직접 다운로드:**
   - https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js
   - 또는: https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js
   - 다운로드한 파일을 `xlsx.full.min.js`로 이름 변경하여 이 폴더에 저장

2. **npm을 통한 다운로드:**
   ```bash
   npm install xlsx
   # node_modules/xlsx/dist/xlsx.full.min.js 파일을 이 폴더에 복사
   ```

3. **curl 명령어 사용:**
   ```bash
   curl -L -o xlsx.full.min.js https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js
   ```

### Option 2: 내부망 전송 (Transfer to Internal Network)

1. 외부 인터넷이 가능한 PC에서 위 방법으로 다운로드
2. USB 드라이브나 승인된 파일 전송 방법을 통해 내부망으로 전송
3. 이 디렉토리(`lib/`)에 `xlsx.full.min.js` 파일 배치

### 파일 구조 확인 (Verify File Structure)

설치 후 다음과 같은 구조가 되어야 합니다:

```
excelauto/
├── index.html
├── app.js
├── lib/
│   ├── xlsx.full.min.js  ← 이 파일이 필요함
│   └── README.md (이 파일)
└── templates/
```

### 라이브러리 정보 (Library Information)

- **이름:** SheetJS Community Edition
- **버전:** 0.18.5 또는 0.20.0 (호환됨)
- **라이센스:** Apache 2.0
- **용도:** Excel 파일 읽기/쓰기
- **크기:** 약 800KB ~ 1.2MB
- **공식 사이트:** https://sheetjs.com/

### 대체 방법 (Alternative Method)

SheetJS 라이브러리를 다운로드할 수 없는 경우, CSV 형식의 데이터로 작업하거나 
다른 오픈소스 Excel 처리 라이브러리를 사용할 수 있습니다.

## 문제 해결 (Troubleshooting)

### 라이브러리가 로드되지 않을 때:
1. `xlsx.full.min.js` 파일이 `lib/` 폴더에 있는지 확인
2. 파일명이 정확히 일치하는지 확인 (대소문자 구분)
3. 브라우저 콘솔(F12)에서 오류 메시지 확인
4. 파일이 손상되지 않았는지 확인 (파일 크기가 0이 아닌지)

### 보안 정책으로 인한 차단:
- 일부 내부망 환경에서는 JavaScript 파일 실행이 제한될 수 있음
- IT 관리자에게 문의하여 로컬 파일 실행 권한 확인
