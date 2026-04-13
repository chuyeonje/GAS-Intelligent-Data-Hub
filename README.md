Master-Table Driven Framework
Token-Optimized Action AI on Google Apps Script

(시트 :  https://docs.google.com/spreadsheets/d/1Emh7n5RNDgkIDQtlHE186rtQdcoJn8GydOaOuVVv3Tw/edit?usp=drive_link )

본 프로젝트는 Google Sheets와 Google Apps Script(GAS)를 기반으로 구축된 경량형 액션 AI 자동화 프레임워크입니다.
단순한 스프레드시트 시스템이 아니라, 관리 테이블(Master Table)을 중심으로 동작하는 구조화된 실행 엔진입니다.

🧠 핵심 설계 철학

이 시스템은 다음 3가지 원칙을 기반으로 설계되었습니다.

Spreadsheet = Database Layer
GAS = Execution Layer
Master Table = System Kernel

즉, 모든 동작은 “관리 테이블”을 기준으로 해석되고 실행됩니다.

🚀 주요 기술 특징
1. In-Memory SQL Engine (AlaSQL 기반)
Google Sheets 데이터를 클라이언트 메모리로 로드
AlaSQL을 이용해 SQL 문법으로 데이터 처리
SELECT / INSERT / UPDATE / DELETE 전부 지원
시트 직접 접근 대신 메모리 기반 배치 처리 구조

➡️ 결과: 반복적인 GAS 호출 제거 및 처리 비용 감소

2. Zero-Dependency Lifecycle Control (Zombie-Free UI)
requestAnimationFrame 기반 렌더링 감지
GAS Dialog 종료 시 백그라운드 프로세스 자동 종료
오디오 / 세션 / 타이머 완전 해제

➡️ 결과: “좀비 프로세스” 없이 안정적인 lifecycle 관리

3. Chunk-Based Media Streaming
대용량 오디오 파일을 Base64 chunk로 분할 저장
병렬 Promise.all() 기반 다운로드
순차 다운로드 대비 초기 latency 감소

➡️ 결과: GAS 환경에서도 스트리밍 유사 경험 구현

4. Master Table Driven Architecture
모든 시트 구조는 “관리 테이블”이 정의
컬럼 구조 / 정규식 / 함수까지 중앙 제어
런타임에 구조 변경 없이 데이터 규칙 적용

➡️ 결과: “코드가 아니라 테이블이 시스템을 정의”

5. Slack-Based Remote Command Interface
Slack Webhook 기반 외부 제어 시스템
/command [sheet] action 구조로 명령 처리
시트 단위 컨텍스트 라우팅 지원

➡️ 결과: 외부에서 시스템 직접 제어 가능

6. Hybrid Data Indexing (RAG-lite 구조)
Drive / Sheet / Log 데이터를 단일 인덱스로 추상화
Keyword pre-filtering 기반 데이터 축소
LLM 전달 컨텍스트 최소화

➡️ 결과: API 비용 없이 유사 RAG 구조 구현

🤖 Action-Oriented AI Agent Model

본 시스템의 AI는 “응답형 챗봇”이 아니라 실행형 에이전트입니다.

질문 응답 ❌
작업 수행 ✅

지원 작업:

시트 CRUD
드라이브 파일 관리
일정 자동 입력
외부 데이터 수집 후 반영

🏗 시스템 구조
Top-Down Control Model
모든 시트는 Master Table 정의를 따름
신규 데이터는 자동 검증 (Regex 기반)
수식 자동 주입 지원
Hybrid Query Layer
복잡한 조회: AlaSQL (client side)
대규모 필터링: Google Sheets native function
실행 레이어 분리로 성능 최적화

⚙️ 설치 구조
Google-Drive-Root/
 └── 시트폴더/
      ├── MUSIC/
      ├── MUSICLIST/
      ├── AI/
      ├── ORC/
      └── [Master Control Sheet]

🔧 Setup
Drive 폴더 구조 업로드
Master Sheet 복사
GAS Web App 배포
환경 변수 설정
GEMINI_API_KEY
SLACK_WEBHOOK_URL
FOLDER_ID
Master Table Sync 실행

⚠️ 주의사항
구조 변경은 반드시 Master Table을 통해 수행
시트 직접 수정은 시스템 불일치 발생 가능
대규모 작업 전 반드시 백업 권장

🧩 핵심 요약
이 시스템은 “스프레드시트 기반 자동화 툴”이 아니라
스프레드시트를 데이터베이스로 재해석한 실행형 AI OS 구조입니다.
