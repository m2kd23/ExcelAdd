# ⚡ 내부망용 오프라인 엑셀 취합 도구 (Excel Merge Tool)

<p align="left">
  <img src="https://img.shields.io/badge/Environment-100%25%20Offline%20(망분리)-success?style=flat-square&logo=shield" alt="Offline">
  <img src="https://img.shields.io/badge/Installation-Zero%20Install-blue?style=flat-square" alt="Zero Install">
  <img src="https://img.shields.io/badge/Server-Zero%20Server%20(Client--Only)-orange?style=flat-square" alt="No Server">
  <img src="https://img.shields.io/badge/Data%20Privacy-No%20Data%20Leakage-brightgreen?style=flat-square" alt="Privacy">
  <img src="https://img.shields.io/badge/License-MIT-lightgrey?style=flat-square" alt="License">
</p>

> **"설치 No, 서버 No, 인터넷 No!"**  
> 공공기관 및 기업 내부망 PC에서 `index.html` 더블 클릭 한 번으로 여러 부서의 엑셀 파일을 안전하고 신속하게 취합하는 **100% 브라우저 메모리 기반 스마트 취합 도구**입니다.

---

## 🌟 핵심 특징 (Key Highlights)

| 기능 | 설명 |
| :--- | :--- |
| 🔒 **100% 오프라인 & 완벽 보안** | 외부 서버 통신(`fetch`, `API`, `CDN`)이 일체 없으며, 사용자 PC 메모리(RAM) 내부에서만 연산되어 데이터 외부 유출이 원천 차단됩니다. |
| 🚀 **Zero-Install (무설치)** | Python, WAS, Node.js, DB 설치가 전혀 필요 없으며, 브라우저(Edge, Chrome, Whale)만 있으면 즉시 실행됩니다. |
| 📋 **2대 취합 모드 지원** | **`[전체 취합]`**(단순 행 합치기) 및 **`[취합양식에서 작성 부분만 취합]`**(지능형 양식 매칭 & 제자리 행 삽입)을 모두 지원합니다. |
| 🧠 **지능형 서식 분석 (Smart Engine)** | 문서 상단 배너나 공백 줄이 있어도 실제 제목줄을 스스로 감지하며, 중간에 임의로 추가된 행도 원래 문맥 위치에 **1행 전체(Full-Row)**로 정확히 삽입합니다. |
| 🛡️ **비파괴 무결성 보장** | 원본 엑셀 파일은 절대 훼손하지 않으며, 항상 새로운 취합 결과 엑셀로 자동 다운로드됩니다. |

---

## 🏗️ 시스템 아키텍처 (System Architecture)

별도의 백엔드 없이 클라이언트 브라우저 런타임에서 완결되는 **Zero-Server Client-Side 구조**입니다.

```mermaid
flowchart LR
    subgraph ClientPC [" 사용자 PC (내부망 / 망분리 환경) "]
        Folder[("📂 취합 대상 폴더\n(.xlsx / .xls)")]
        
        subgraph Engine [" 브라우저 런타임 (Memory) "]
            UI["🖥️ index.html\n(사용자 인터페이스)"]
            Core["⚙️ 취합 컨트롤러\n(Smart Engine)"]
            SheetJS["📦 SheetJS 코어\n(xlsx.full.min.js)"]
        end
        
        Result[("📊 최종 취합 엑셀\n(자동 다운로드)")]
    end

    ExternalNetwork["❌ 외부 네트워크 (차단됨)"] -.x Engine

    Folder -->|FileReader API| Core
    Core <--> SheetJS
    Core -->|Blob 생성| Result
    UI <--> Core

    style ClientPC fill:#f8f9fa,stroke:#1971c2,stroke-width:2px
    style Engine fill:#e7f5ff,stroke:#1971c2,stroke-width:1px
    style ExternalNetwork fill:#fff5f5,stroke:#c92a2a,stroke-width:1px
```

---

## 🎯 주요 취합 방식 비교

### 1. [전체 취합] (구: 단순 행 합치기)
* **용도**: 여러 부서에서 제출한 목록성 데이터(1팀, 2팀, 3팀 설문 결과 등)를 하나의 엑셀에 아래로 쭉 이어 붙일 때
* **동작 원리**:
  1. 각 파일의 실제 데이터 제목줄(Header)을 자동 감지 (상단 안내문/배너 자동 제외)
  2. 최상단에 제목줄을 1회만 표시하고, 모든 파일의 데이터 행을 아래로 연속 결합
  3. 컬럼 순서가 서로 달라도 컬럼명을 기준으로 일치하는 위치에 자동 정렬

```text
[1팀.xlsx] (A, B행)  ──┐
[2팀.xlsx] (C, D행)  ──┼──> [전체취합결과.xlsx] (A, B, C, D, E, F행 통합 목록)
[3팀.xlsx] (E, F행)  ──┘
```

---

### 2. [취합양식에서 작성 부분만 취합] (구: 양식 엑셀 취합 / 기준값 매칭)
* **용도**: 배포된 서식(예: 담당자조사, 홈페이지 점검표 등)에 각 부서가 담당자명/수정내용을 채워 넣고, 필요에 따라 중간에 신규 행을 추가해 회신한 경우
* **동작 원리**:
  1. **스마트 뼈대 분석**: 마스터 파일을 따로 지정하지 않아도 폴더 내 서식을 자동 파악
  2. **빈 셀 쏙쏙 채우기**: 기존 서식 틀은 그대로 보존하고, 부서별 작성 내용만 빈 셀에 안전하게 병합
  3. **문맥 기반 제자리 행 삽입 (In-Place Insertion)**:
     - 부서에서 중간에 행을 추가한 경우, A~F열 전체 1행(Full-Row)을 해당 부서의 작성 위치 바로 아래에 쏙 끼워 넣어 **상하위 메뉴 줄 맞춤이 단 1칸도 밀리지 않도록 완벽 정렬**

```text
[배포 원본 서식] ─── (각 부서 배포) ───> [인사팀.xlsx (담당자 입력 + 15행 신규행 추가)]
                                  [총무팀.xlsx (담당자 입력)]
                                            │
                                            ▼ (원클릭 스마트 취합)
[최종 결과.xlsx] ➔ 인사팀 담당자 채워짐 + 15행에 신규행 쏙 들어감 + 총무팀 담당자 채워짐 (줄 어긋남 0%)
```

---

## 🖥️ 사용 방법 (Step-by-Step)

```text
1. index.html 더블 클릭 실행
   └── 웹 브라우저(Edge, Chrome, Whale)에서 즉시 열립니다.

2. 취합 방식 선택
   ├── [● 전체 취합] : 여러 엑셀의 모든 행을 하나로 합칠 때 (기본)
   └── [○ 취합양식에서 작성 부분만 취합] : 배포 양식의 작성 내용 및 추가 행을 맞출 때

3. '취합할 폴더' 선택
   └── [폴더 선택] 버튼을 클릭하여 취합 대상 엑셀들이 들어있는 폴더를 선택합니다.
       (※ 임시 파일(~$...) 및 비엑셀 파일은 프로그램이 자동 제외)

4. [취합 시작] 클릭
   └── 1초 만에 취합이 완료되며 브라우저 다운로드 폴더에 결과 엑셀이 자동 저장됩니다!
```

---

## 📁 디렉토리 구조 (Repository Layout)

```text
ExcelAdd/
  ├── index.html                      # [메인] 웹 어플리케이션 UI & 취합 코어 엔진
  ├── xlsx.full.min.js                # [로컬 라이브러리] 오프라인 SheetJS 엑셀 파서
  ├── README.md                       # [문서] GitHub 프로젝트 안내서
  ├── README.txt                      # [문서] 텍스트 퀵 가이드
  └── docs/
      └── 시스템설계서_및_사용자매뉴얼.md   # [산출물] 시스템 아키텍처, DFD 및 상세 사용자 매뉴얼
```

---

## 🛡️ 보안 및 개인정보 무결성 검증

- [x] **외부 네트워크 통신 0건**: `fetch`, `XMLHttpRequest`, `WebSocket`, 외부 CDN 일체 차단
- [x] **개인정보 하드코딩 0건**: 소스 내 개인정보, API 키, 패스워드 전무
- [x] **영구 흔적 소멸**: `localStorage`, `Cookie` 미사용으로 브라우저 종료 시 메모리 완전 소멸
- [x] **XSS 보안 방어**: `textContent` 기반 안전 렌더링으로 악성 파일명 스크립트 실행 방지

---

## 📄 라이선스 (License)

본 프로젝트는 [MIT License](LICENSE)를 따릅니다. 공공기관 및 기업 사내망에서 자유롭게 복제, 수정 및 배포하여 활용하실 수 있습니다.
