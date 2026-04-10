// #region [선언부] - 프롬프트 원본 100% 유지
const multiAiDataObject ={
    googleCommander:{
        prompt :
            `
# Role
너는 사용자의 요청을 분석하여, 고정된 파이프라인(검색 -> 드라이브 -> 캘린더 -> 메일 -> 시트)에서 [실행할 작업]과 [데이터 상속 여부]를 판독하는 Dependency Analyzer다.

# Output Format
실행이 필요한 구역만 ^ 으로 연결하여 출력하라. 부가 설명은 절대 금지한다.

# 구역별 작성 규칙
1. 검색(i): 
   - 실행 시: i
2. 드라이브(d): 
   - 단독 실행: d_[키워드1]_[키워드2]
   - 검색(i) 데이터 상속 필요: d_[키워드]_r(i)
   - 특정 검색 키워드가 필요 없는 작업: d 또는 d_r(...)   
3. 시트(s):
   - 단독 실행 (키워드 추출): s_[키워드1]_[키워드2]
   - 이전 데이터 상속 필요: s_[키워드]_r(i) 또는 s_[키워드]_r(d) 또는 s_[키워드]_r(i,d)
   - 특정 검색 키워드가 필요 없는 시트 작업: s 또는 s_r(...)
4. 캘린더(c):
   - 단독 실행: c_[키워드1]_[키워드2]
   - 검색(i) 데이터 상속 필요: c_[키워드]_r(i)
   - 특정 검색 키워드가 필요 없는 작업: c 또는 c_r(...)
5. 메일(m): 
   - 단독 실행: m_[키워드]
   - 이전 데이터 상속 필요: m_r(i,c) 등
   - 특정 키워드 불필요: m 또는 m_r(...)

# Examples
- "오늘 삼성전자 주가 검색해서 시트에 적어"
  => i^s_삼성전자_r(i)
- "AI 폴더 지우고, 시트에서 김철수 데이터 찾아"
  => d_AI^s_김철수
- "인터넷에서 최신 IT 뉴스 찾아서 그걸로 드라이브에 문서 하나 만들고, 시트에도 기록해 줘"
  => i^d_r(i)^s_r(i,d)
- "내일 서울 날씨 검색해 줘"
  => i
- "드라이브에 임시 폴더 하나 만들어"
  => d
- "내일 오후 3시에 회의 일정 잡고, 시트에도 기록해 줘"
  => c^s_r(c)
- "오늘 서울 날씨 검색해서 내일 오전 10시 캘린더에 날씨 알림 일정으로 추가해"
  => i^c_r(i)  
- "오늘 삼성전자 주가 검색해서 시트에 적어"
  => i^s_삼성전자_r(i)

  [사용자의 요청]
  - `
    },
    googleSearch:{
        prompt : `
# Role
너는 Google Search 도구를 사용하여 실시간 정보를 수집하고, 불필요한 사족 없이 핵심 데이터만 추출하는 전문 리서처야
# 주의사항
1.무조건 Google 검색을 수행하여 최신 정보를 확인해. (학습된 데이터로 지어내지 마)
2.출력물에는 인사말, 서론, 결론, 출처 표기, 마크다운 기호를 절대 포함하지 마.
3.오직 사용자가 요청한 정보(가사, 데이터 등) 전문만 텍스트로 출력해.
4.검색 결과가 여러 개라면 가장 정확한 하나를 선택하고, 정보가 없다면 "false"라고만 출력해.
# Target Data
사용자의 요청 내용을 분석하여 해당 데이터의 전체 텍스트(Full Text)를 확보해.

[사용자 요청]
- `
    },
    googleSheet:{
        prompt : `
      # Role
      너는 구글 시트 기반의 자율 DB 관리 에이전트다.
      사용자의 요청을 분석하여, 시트를 제어하기 위한 최적의 명령어(Opcode) 조합을 구성한 뒤 반드시 execute_opcodes 도구를 호출하여 응답하라.

      # Strict Rules (최우선 준수 사항)
      1. [작업구역 정의]
          - 2:2 행 데이터는 컬럼명으로 취급한다.
          - 모든 시트의 데이터 시작 위치는 B3부터 시작한다.
          - 1:1,2:2 및 A열은 보호 구역이다 절대로 생성,수정,삭제를 할 수 없다.
          - 컬럼이 없는 열에 데이터는 추가 할 수 없으며 데이터는 컬럼에 맞춰 배치한다.
          - 수평은 하나의 ID로 보고 데이터 추가는 수직으로 이루어진다.
      2. [데이터 무결성] 사용자의 별도의 요청이 없다면 외부에서 검색된 데이터는 임의 요약이나 가공 없이, 원문 100%를 value 파라미터에 할당하라.
      3. [대량 데이터 최적화] 구구단이나 표처럼 2개 이상의 셀을 입력할 때는 절대! 반드시! 무조건! 배열 입력(7번)을 사용하여 명령어 객체(Object) 단일화하라!
      4. [행 단위 연산] 단일 행의 다수 셀을 비우거나 조작할 경우, 개별 셀 제어 대신 행 완전 삭제(8번) 및 '빈 행 삽입(9번)'을 우선 적용하여 연산을 최적화하라.
      5. [마지막 행의 규칙] target,value의 마지막 행이나 빈 시트의 처음 데이터를 추가하는 경우 입력 값에 lt를 붙힌다.
        ( 예 :
         - ltB0 === 마지막 행
         - ltB3 === 마지막 행+3
         - ltB0:ltB3 === 마지막 행:마지막행+3
         )

      # Opcode Reference (도구 파라미터 매핑 규칙)
      execute_opcodes 도구의 commands 배열 내 각 객체는 다음 규칙에 따라 작성한다:
      - 1 단일 입력: target="B3", value="사과"
      - 2 단일 삭제: target="C4", value="" (빈 문자열)
      - 3 단일 이동: target="B3", value="C3" (도착지 셀 주소)
      - 4 범위 입력: target="B3:B5", value="완료"
      - 5 범위 삭제: target="C3:C5", value=""
      - 6 범위 이동: target="B3:B5", value="D3:D5" (도착지 범위)
      - 7 2차원 배열: target="B3", value="2단,2x1=2|3단,3x1=3" (주의: 반드시 쉼표(,)로 열(Column)을 구분하고, 파이프(|)로 행(Row)을 구분하라.)
      - 8 행 삭제: target="ROW5", value=""
      - 9 행 삽입: target="ROW5", value=""

      # Execution
      [사용자 요청], [시트 데이터], [검색한 데이터]를 분석하고 출력물에 # Strict Rules 와 # Opcode Reference 규칙을 교차 검증하라. 검증이 끝나면 텍스트 답변 없이 즉시 도구를 호출하라.

      [사용자 요청]
      -`,
        sheetData :'\n\n[시트 데이터]:\n',
        functionDeclarations : [
            {
                "name": "execute_opcodes",
                "description": "사용자의 요청을 분석하여 구글 시트 관리용 Opcode 명령어 리스트를 실행합니다.",
                "parameters": {
                    "type": "OBJECT",
                    "properties": {
                        "commands": {
                            "type": "ARRAY",
                            "description": "실행할 명령어들의 순차적 리스트",
                            "items": {
                                "type": "OBJECT",
                                "properties": {
                                    "opcode": {
                                        "type": "STRING",
                                        "description": "명령어 번호 (1:단일생성, 2:단일삭제, 3:단일이동, 4:범위수정, 5:범위삭제, 6:범위이동, 7:2차원배열, 8:행삭제, 9:행삽입)",
                                        "enum": ["1", "2", "3", "4", "5", "6", "7", "8", "9"]
                                    },
                                    "target": {
                                        "type": "STRING",
                                        "description": "대상 셀 범위 또는 행 번호 (예: 'B3', 'A1:A5', 'ROW5')"
                                    },
                                    "value": {
                                        "type": "STRING",
                                        "description": "입력값, 목적지 주소 또는 배열 데이터 (삭제 명령의 경우 빈 값)"
                                    }
                                },
                                "required": ["opcode", "target"]
                            }
                        }
                    },
                    "required": ["commands"]
                }
            }
        ]
    },
    googleDrive:{
        prompt : `
      # Role
      너는 구글 드라이브 자율 관리 에이전트다. 
      사용자의 자연어 요청을 분석하여, 오직 제공된 execute_opcodes 도구를 호출하는 방식으로만 응답하라. 일반 텍스트 설명은 절대 출력하지 마라.

      # Rules
      1. [ID 기반 매핑] 드라이브 내 파일/폴더 이름은 중복될 수 있으므로, 반드시 [드라이브 데이터]를 대조하여 고유한 "ID"를 추출한 뒤 도구의 파라미터로 사용하라.
      2. [단일 책임의 원칙] 여러 파일/폴더를 제어해야 할 경우, commands 배열 내에 개별 명령어 객체(Object)를 여러 개 생성하여 순차적으로 배열하라.
      3. [안전 삭제] 삭제(4) 명령어는 영구 삭제가 아닌 '휴지통 이동'을 의미한다.
      4. [데이터 요청] 사용자의 요청으로 검색 데이터가 필요한 경우 [검색 데이터]를 사용한다.

      # Opcode Reference
      도구 호출 시 commands 배열 내 각 객체는 다음 규칙에 따라 작성한다:
      - 1 새 폴더 생성: target="부모 폴더 ID", value="새 폴더 이름"
      - 2 이동: target="이동할 파일/폴더 ID", value="목적지 폴더 ID"
      - 3 이름 변경: target="대상 파일/폴더 ID", value="바꿀 이름"
      - 4 휴지통 이동(삭제): target="삭제할 파일/폴더 ID", value="" (빈 문자열)
      - 5 복사: target="원본 파일 ID", value="목적지 폴더 ID"
      - 6 텍스트 파일 생성: target="부모 폴더 ID", value="파일명.txt|파일내용" (파이프 '|' 기호로 파일명과 내용을 구분)

      # Execution
      [사용자 요청]을 검증하여 타겟의 정확한 ID를 파악하라. 검증이 끝나면 텍스트 답변 없이 즉시 도구를 호출하라.

      [사용자 요청]
      - `
        ,
        functionDeclarations :
            [
                {
                    "name": "execute_opcodes",
                    "description": "사용자의 요청을 분석하여 구글 드라이브 관리용 Opcode 명령어 리스트를 실행합니다.",
                    "parameters": {
                        "type": "OBJECT",
                        "properties": {
                            "commands": {
                                "type": "ARRAY",
                                "description": "실행할 명령어들의 순차적 리스트",
                                "items": {
                                    "type": "OBJECT",
                                    "properties": {
                                        "opcode": {
                                            "type": "STRING",
                                            "description": "명령어 번호 (1:폴더생성, 2:파일/폴더이동, 3:이름 변경, 4:휴지통 이동, 5:파일 복사, 6:텍스트 파일 생성)",
                                            "enum": ["1", "2", "3", "4", "5", "6"]
                                        },
                                        "target": {
                                            "type": "STRING",
                                            "description": "타겟ID 또는 부모폴더ID"
                                        },
                                        "value": {
                                            "type": "STRING",
                                            "description": "목적지ID 또는 이름(삭제 명령의 경우 빈 값)"
                                        }
                                    },
                                    "required": ["opcode", "target"]
                                }
                            }
                        },
                        "required": ["commands"]
                    }
                }
            ]
    },
    googleCalendar:{
        prompt : `
      # Role
      너는 구글 캘린더 자율 관리 에이전트다.
      사용자의 자연어 요청과 현재 시간 정보를 분석하여, 오직 제공된 execute_opcodes 도구를 호출하는 방식으로만 응답하라. 일반 텍스트 설명은 절대 출력하지 마라.

      # Rules
      1. [시간 컨텍스트 절대 규칙] 모든 시간은 한국(KST) 기준이다. 시스템 현재 시간을 참고하여 반드시 'YYYY-MM-DDTHH:mm:ss' 포맷으로 작성하라.
         (예: 4월 8일 낮 11시 -> 2026-04-08T11:00:00)
      2. [기본 소요 시간] 약속의 종료 시간이 명시되지 않았다면, 시작 시간으로부터 기본 '1시간' 뒤로 종료 일시를 설정하라.
      3. [단일 책임의 원칙] 여러 일정을 제어하거나 다중 작업을 해야 할 경우, commands 배열 내에 개별 명령어 객체를 순차적으로 배열하라.
      4. [구분자 규칙] 다중 파라미터가 필요한 경우 반드시 파이프(|) 기호로 구분하여 value에 할당하라. 빈 값은 생략하되 파이프 기호는 유지한다 (예: 제목|시작시간|종료시간||장소).
      5. [데이터 연계] 사용자의 요청이 캘린더 일정 '검색/조회'인 경우, 4번 명령어를 사용해 데이터를 추출한다. 이 데이터는 파이프라인의 다음 단계(시트/드라이브/메일)로 전달된다.

      # Opcode Reference
      도구 호출 시 commands 배열 내 각 객체는 다음 규칙에 따라 작성한다:
      - 1 일정 생성: target="PRIMARY" (또는 특정 캘린더ID), value="제목|시작일시(ISO)|종료일시(ISO)|설명|위치"
      - 2 일정 삭제: target="삭제할 이벤트 ID", value="" (빈 문자열)
      - 3 일정 제목 변경: target="변경할 이벤트 ID", value="새로운 제목"
      - 4 일정 검색/조회: target="PRIMARY", value="조회시작일시(ISO)|조회종료일시(ISO)|검색키워드(선택)" (기간 내 일정을 배열로 반환하기 위함)

      # Execution
      [사용자 요청] 및 [현재 시간]을 분석하여 정확한 타겟과 값을 파악하라. 검증이 끝나면 텍스트 답변 없이 즉시 도구를 호출하라.

      [사용자 요청]
      -   
    `,
        functionDeclarations :
            [
                {
                    "name": "execute_opcodes",
                    "description": "사용자의 요청을 분석하여 구글 캘린더 관리용 Opcode 명령어 리스트를 실행합니다.",
                    "parameters": {
                        "type": "OBJECT",
                        "properties": {
                            "commands": {
                                "type": "ARRAY",
                                "description": "실행할 명령어들의 순차적 리스트",
                                "items": {
                                    "type": "OBJECT",
                                    "properties": {
                                        "opcode": {
                                            "type": "STRING",
                                            "description": "명령어 번호 (1:일정생성, 2:일정삭제, 3:제목변경, 4:일정검색)",
                                            "enum": ["1", "2", "3", "4"]
                                        },
                                        "target": {
                                            "type": "STRING",
                                            "description": "캘린더 ID (기본값: 'PRIMARY') 또는 대상 이벤트 ID"
                                        },
                                        "value": {
                                            "type": "STRING",
                                            "description": "파이프(|)로 구분된 매개변수 값 (생성: 제목|시작|종료|설명|위치, 검색: 시작|종료|키워드)"
                                        }
                                    },
                                    "required": ["opcode", "target"]
                                }
                            }
                        },
                        "required": ["commands"]
                    }
                }
            ],
    },
    googleGmail: {
        prompt: `
      # Role
      너는 구글 Gmail 자율 관리 에이전트다.
      사용자의 요청과 전달받은 데이터를 분석하여, 오직 execute_opcodes 도구를 호출해 응답하라.

      # Rules
      1. [명령어 규칙] 메일 전송 시 제목과 본문을 파이프(|)로 구분한다. 수신자가 명시되지 않았다면 '나에게 쓰기(Session.getActiveUser().getEmail())'로 간주하여 처리한다.
      2. [데이터 연계] 검색(i), 드라이브(d), 캘린더(c), 시트(s)에서 상속받은 데이터가 있다면, 이를 요약하거나 원문 그대로 메일 본문(value)에 자연스럽게 포함시켜라.

      # Opcode Reference
      - 1 메일 발송: target="수신자 이메일 (없으면 me)", value="메일 제목|메일 본문(상속 데이터 포함)"
      - 2 최근 메일 읽기: target="검색 쿼리 (예: is:unread, from:boss@gmail.com)", value="읽어올 개수(숫자)"

      # Execution
      [사용자 요청] 및 [상속 데이터]를 분석하여 도구를 호출하라.

      [사용자 요청]
      - `,
        functionDeclarations: [
            {
                "name": "execute_opcodes",
                "description": "구글 Gmail 제어용 Opcode 실행",
                "parameters": {
                    "type": "OBJECT",
                    "properties": {
                        "commands": {
                            "type": "ARRAY",
                            "items": {
                                "type": "OBJECT",
                                "properties": {
                                    "opcode": { "type": "STRING", "enum": ["1", "2"] },
                                    "target": { "type": "STRING", "description": "이메일 주소 또는 검색 쿼리" },
                                    "value": { "type": "STRING", "description": "제목|본문 또는 가져올 메일 개수" }
                                },
                                "required": ["opcode", "target"]
                            }
                        }
                    },
                    "required": ["commands"]
                }
            }
        ]
    }
};
// #endregion

/**
 * ============================================================================
 * [진입점 및 헬퍼 함수]
 * ============================================================================
 */
function listAvailableModels() {
    const props = PropertiesService.getScriptProperties();
    const apiKey = props.getProperty('GEMINI_API_KEY');
    const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`;

    const response = UrlFetchApp.fetch(url, { "muteHttpExceptions": true });
    const json = JSON.parse(response.getContentText());

    if (json.models) {
        json.models.forEach(model => Logger.log(`모델명: ${model.name} | 설명: ${model.description}`));
    } else {
        Logger.log("모델 목록 가져오기 실패: " + response.getContentText());
    }
}

function testfunctionAi(){
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const adminSheet = ss.getSheetByName('테스트테이블');
    sttAi(adminSheet, "캘린더에 내일 낮 11시에 약속 잡아줘");
}

function sttAi(sheet, text) {
    const rawContext = autoSheetFilter(sheet);
    Logger.log("입력 자연어 : " + text);

    // 파이프라인 매니저로 캡슐화된 로직 실행
    const resultMsg = AIPipelineManager.execute(text, rawContext, sheet);
    Logger.log("최종 수행 결과 : " + resultMsg);

    // 프론트엔드의 STT 텍스트 증발 방지
    if (resultMsg !== "AI 로직이 정상적으로 수행되었습니다.") {
        return { success: false, error: resultMsg };
    }
    return { success: true };
}

function autoSheetFilter(sheet) {
    const lastNumber = Number(sheet.getRange("B1").getValue()) + 2;
    const lastColNum = sheet.getLastColumn();
    if (lastColNum === 0) return "A";

    const lastColLetter = sheet.getRange(1, lastColNum).getA1Notation().replace(/\d+/, "");
    const range = sheet.getRange("A2:" + lastColLetter + lastNumber);

    return range.getDisplayValues()
        .filter(row => row.some(cell => cell !== ""))
        .map(row => row.join(" | "))
        .join("\n");
}

/**
 * ============================================================================
 * [AI API 통신 전담 객체]
 * ============================================================================
 */
const GeminiAPI = {
    fetch: function(prompt, mode = 0) {
        const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
        const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite-preview:generateContent?key=${apiKey}`;

        let payload = {
            "contents": [{ "parts": [{ "text": prompt }] }],
            "tools": [],
            "generationConfig": { "temperature": 0.2, "maxOutputTokens": 8192 }
        };

        switch (mode) {
            case 1: // Router
                payload.generationConfig.temperature = 0.0;
                break;
            case 2: // Search
                payload.tools.push({ "googleSearch": {} });
                break;
            case 3: // Sheet
                if (multiAiDataObject.googleSheet.functionDeclarations) payload.tools.push({ "functionDeclarations": multiAiDataObject.googleSheet.functionDeclarations });
                payload.generationConfig.temperature = 0.0;
                break;
            case 4: // Drive
                if (multiAiDataObject.googleDrive.functionDeclarations) payload.tools.push({ "functionDeclarations": multiAiDataObject.googleDrive.functionDeclarations });
                payload.generationConfig.temperature = 0.0;
                break;
            case 5: // Calendar
                if (multiAiDataObject.googleCalendar.functionDeclarations) payload.tools.push({ "functionDeclarations": multiAiDataObject.googleCalendar.functionDeclarations });
                payload.generationConfig.temperature = 0.0;
                break;
            case 6: // Mail
                if (multiAiDataObject.googleGmail.functionDeclarations) payload.tools.push({ "functionDeclarations": multiAiDataObject.googleGmail.functionDeclarations });
                payload.generationConfig.temperature = 0.0;
                break;
        }

        return this.executeWithRetry(apiUrl, payload, mode);
    },

    executeWithRetry: function(apiUrl, payload, mode) {
        const options = {
            "method": "post",
            "contentType": "application/json",
            "payload": JSON.stringify(payload),
            "muteHttpExceptions": true
        };

        let maxRetries = 3;
        let delay = 3000;

        for (let attempt = 0; attempt < maxRetries; attempt++) {
            try {
                const response = UrlFetchApp.fetch(apiUrl, options);
                if (response.getResponseCode() !== 200) {
                    throw new Error(`HTTP ${response.getResponseCode()}: ${response.getContentText()}`);
                }

                const json = JSON.parse(response.getContentText());
                if (json.candidates && json.candidates[0]?.content?.parts[0]) {
                    const parts = json.candidates[0].content.parts[0];
                    return parts.functionCall ? JSON.stringify(parts.functionCall.args) : parts.text.trim();
                }
                throw new Error("유효한 데이터 없음");
            } catch (e) {
                Logger.log(`[API 실패 - 시도 ${attempt + 1}/${maxRetries}] 에러: ${e.message}`);
                if (attempt < maxRetries - 1) {
                    Utilities.sleep(delay); delay *= 2;
                } else {
                    return (mode === 2) ? "false" : `Error: API 실패 - ${e.message}`;
                }
            }
        }
        return "false";
    }
};

/**
 * ============================================================================
 * [파이프라인 오케스트레이터]
 * ============================================================================
 */
const AIPipelineManager = {
    execute: function(command, rawContext, sheet) {
        Logger.log("command: " + command);
        const tablesName = sGetSheetNames().join(",");
        const routeStr = GeminiAPI.fetch(multiAiDataObject.googleCommander.prompt + command +`\n\n[시트명]:\n${tablesName}` , 1);
        Logger.log("[라우터 판독 결과]: " + routeStr);

        if (routeStr === "false" || routeStr.startsWith("Error")) return "라우터 판독 실패";

        const tasks = routeStr.split('^');
        let ctx = { search: "", drive: "", calendar: "", mail: "", sheetRaw: rawContext };

        if (!this.runSearch(tasks, command, ctx)) return "검색 중단";
        if (!this.runDrive(tasks, command, ctx)) return "드라이브 제어 중단";
        if (!this.runCalendar(tasks, command, ctx)) return "캘린더 제어 중단";
        if (!this.runMail(tasks, command, ctx)) return "메일 제어 중단";

        // 시트 제어는 파이프라인의 최종 목적지이므로 에러 시에도 결과를 반환하도록 처리
        try {
            this.runSheet(tasks, command, ctx, sheet);
        } catch (e) {
            return `시트 제어 중단: ${e.message}`;
        }

        return "AI 로직이 정상적으로 수행되었습니다.";
    },

    buildInherit: function(taskStr, ctx) {
        let prompt = "";
        const match = taskStr.match(/r\(([^)]+)\)/);
        if (match) {
            const inherits = match[1];
            if (inherits.includes('i') && ctx.search) prompt += `\n\n[검색 데이터]:\n${ctx.search}`;
            if (inherits.includes('d') && ctx.drive) prompt += `\n\n[드라이브 데이터]:\n${ctx.drive}`;
            if (inherits.includes('c') && ctx.calendar) prompt += `\n\n[캘린더 데이터]:\n${ctx.calendar}`;
            if (inherits.includes('m') && ctx.mail) prompt += `\n\n[메일 데이터]:\n${ctx.mail}`;
        }
        return prompt;
    },

    runSearch: function(tasks, command, ctx) {
        if (!tasks.some(t => t.startsWith('i'))) return true;
        Logger.log("--> [i] 검색 구역 실행");

        ctx.search = GeminiAPI.fetch(multiAiDataObject.googleSearch.prompt + command, 2);
        if (ctx.search === "false" || ctx.search.startsWith("Error")) {
            Logger.log("검색 실패"); return false;
        }
        return true;
    },

    runDrive: function(tasks, command, ctx) {
        const task = tasks.find(t => t.startsWith('d'));
        if (!task) return true;
        Logger.log("--> [d] 드라이브 구역 실행");

        const keywords = task.split('_').filter(p => p !== 'd' && !p.startsWith('r('));
        let finalSheetContext = "";

        if (keywords.length > 0) {
            try {
                const rows = getJsonDataAsText("AI폴더").split('\n');
                const filteredRows = rows.filter(row => keywords.some(k => row.includes(k)));
                finalSheetContext = filteredRows.length > 0 ? filteredRows.join('\n') : "해당 키워드의 데이터가 없습니다.";
            } catch (e) {
                Logger.log("드라이브 DB 로드 실패: " + e.message);
            }
        }

        let driveInput = multiAiDataObject.googleDrive.prompt + command;
        if (finalSheetContext) driveInput += `\n\n[드라이브 데이터]:\n${finalSheetContext}\n`;
        driveInput += this.buildInherit(task, ctx);

        const dResStr = GeminiAPI.fetch(driveInput, 4);
        if (dResStr === "false" || dResStr.startsWith("Error")) return false;

        const results = processAgentResponse(dResStr, WorkspaceActions.Drive);
        if (results.length > 0) ctx.drive = JSON.stringify(results);
        return true;
    },

    runCalendar: function(tasks, command, ctx) {
        const task = tasks.find(t => t.startsWith('c'));
        if (!task) return true;
        Logger.log("--> [c] 캘린더 구역 실행");

        let calInput = multiAiDataObject.googleCalendar.prompt + command;
        const currentTime = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd'T'HH:mm:ssXXX");
        calInput += `\n\n[시스템 현재 시간]: ${currentTime}\n`;

        calInput += this.buildInherit(task, ctx);

        const cResStr = GeminiAPI.fetch(calInput, 5);
        if (cResStr === "false" || cResStr.startsWith("Error")) return false;

        const results = processAgentResponse(cResStr, WorkspaceActions.Calendar);
        if (results.length > 0) ctx.calendar = results.join('\n');
        return true;
    },

    runMail: function(tasks, command, ctx) {
        const task = tasks.find(t => t.startsWith('m'));
        if (!task) return true;
        Logger.log("--> [m] 메일 구역 실행");

        let mailInput = multiAiDataObject.googleGmail.prompt + command + this.buildInherit(task, ctx);
        const mResStr = GeminiAPI.fetch(mailInput, 6);
        if (mResStr === "false" || mResStr.startsWith("Error")) return false;

        const results = processAgentResponse(mResStr, WorkspaceActions.Gmail);
        if (results.length > 0) ctx.mail = results.join('\n');
        return true;
    },

    runSheet: function(tasks, command, ctx, sheet) {
        const task = tasks.find(t => t.startsWith('s'));
        if (!task) return;
        Logger.log("--> [s] 시트 구역 실행");

        // 1. 라우터가 뽑아준 키워드 추출 (예: ["테스트테이블", "구구단"])
        const keywords = task.split('_').filter(p => p !== 's' && !p.startsWith('r('));

        // [수정 포인트 1] 여기서 rows를 반드시 미리 선언하고 쪼개야 합니다!
        const rows = ctx.sheetRaw.split('\n');
        let finalContext = ctx.sheetRaw;

        // 헤더와 꼬리표만 추출하는 내부 함수 (토큰 최적화용)
        const getCompactContext = () => {
            if (rows.length > 4) {
                return rows[0] + "\n...\n(중략)\n...\n" + rows.slice(-3).join('\n');
            }
            return ctx.sheetRaw;
        };

        if (keywords.length > 0) {
            // 키워드가 있는 경우: 일단 시트에서 해당 글자가 있는지 검색해봄
            const filteredRows = rows.filter(row => keywords.some(k => row.includes(k)));

            if (filteredRows.length > 0) {
                // [검색 성공] 기존 데이터를 찾거나 수정하는 명령 (관련 행만 보냄)
                finalContext = filteredRows.join('\n');
            } else {
                Logger.log("키워드 매칭 데이터 없음 -> 신규 데이터 생성으로 간주하여 헤더 전송");
                finalContext = getCompactContext();
            }
        } else {
            // 키워드가 아예 없는 경우 (단순 추가)
            finalContext = getCompactContext();
        }

        let sheetPrompt = multiAiDataObject.googleSheet.prompt + command + `\n\n[시트 데이터]:\n${finalContext}\n`;
        sheetPrompt += this.buildInherit(task, ctx);

        const sResStr = GeminiAPI.fetch(sheetPrompt, 3);
        if (sResStr !== "false" && !sResStr.startsWith("Error")) {
            processAgentResponse(sResStr, WorkspaceActions.Sheet, sheet);
        } else {
            throw new Error("시트 명령어 생성 실패");
        }
    }
};

/**
 * ============================================================================
 * [Action Maps] - 거대한 Switch 문을 대체하는 OCP 기반의 명령어 매핑 객체
 * ============================================================================
 */
const WorkspaceActions = {
    Sheet: {
        "1": (sheet, tgt, val) => sheet.getRange(tgt).setValue(val),
        "2": (sheet, tgt) => sheet.getRange(tgt).clearContent(),
        "3": (sheet, tgt, val) => sheet.getRange(tgt).moveTo(sheet.getRange(val)),
        "4": (sheet, tgt, val) => sheet.getRange(tgt).setValue(val),
        "5": (sheet, tgt) => sheet.getRange(tgt).clearContent(),
        "6": (sheet, tgt, val) => {
            const src = sheet.getRange(tgt);
            const vals = src.getValues();
            const targetStartCell = sheet.getRange(val).getCell(1, 1);
            targetStartCell.offset(0, 0, vals.length, vals[0].length).setValues(vals);
            src.clearContent();
        },
        "7": (sheet, tgt, val) => handleOpcode7(sheet, tgt, val),
        "8": (sheet, tgt) => handleOpcode8(sheet, tgt),
        "9": (sheet, tgt) => handleOpcode9(sheet, tgt)
    },

    Drive: {
        "1": (cmd) => { (cmd.target === "ROOT" ? DriveApp.getRootFolder() : DriveApp.getFolderById(cmd.target)).createFolder(cmd.value); },
        "2": (cmd) => { getDriveItemById(cmd.target).moveTo(DriveApp.getFolderById(cmd.value)); },
        "3": (cmd) => { getDriveItemById(cmd.target).setName(cmd.value); },
        "4": (cmd) => { getDriveItemById(cmd.target).setTrashed(true); },
        "5": (cmd) => { DriveApp.getFileById(cmd.target).makeCopy(DriveApp.getFolderById(cmd.value)); },
        "6": (cmd) => {
            const splitIdx = cmd.value.indexOf('|');
            if (splitIdx === -1) throw new Error("파일 형식 오류 (파이프 없음)");
            const fName = cmd.value.substring(0, splitIdx);
            const fContent = cmd.value.substring(splitIdx + 1);
            (cmd.target === "ROOT" ? DriveApp.getRootFolder() : DriveApp.getFolderById(cmd.target)).createFile(fName, fContent, MimeType.PLAIN_TEXT);
        }
    },

    Calendar: {
        // "1"번: 이벤트 생성 시
        "1": (cmd) => {
            const parts = cmd.value.split('|');
            const cal = cmd.target === "PRIMARY" ? CalendarApp.getDefaultCalendar() : CalendarApp.getCalendarById(cmd.target);

            const startTime = Utils.getAdjustedDate(parts[1]);
            const endTime = Utils.getAdjustedDate(parts[2]);

            return cal.createEvent(parts[0], startTime, endTime, {
                description: parts[3] || "",
                location: parts[4] || ""
            }).getId();
        },
        "2": (cmd, cal) => { cal.getEventById(cmd.target).deleteEvent(); },
        "3": (cmd, cal) => { cal.getEventById(cmd.target).setTitle(cmd.value); },
        // "4"번: 이벤트 조회 시
        "4": (cmd, cal) => {
            const parts = cmd.value.split('|');
            const startTime = Utils.getAdjustedDate(parts[0]);
            const endTime = Utils.getAdjustedDate(parts[1]);
            const opts = parts[2] ? { search: parts[2] } : {};

            return cal.getEvents(startTime, endTime, opts).map(e => ({
                id: e.getId(),
                title: e.getTitle(),
                startTime: e.getStartTime().toISOString(),
                endTime: e.getEndTime().toISOString()
            }));
        }
    },

    Gmail: {
        "1": (cmd) => {
            const targetStr = String(cmd.target || "").trim().toLowerCase();
            const target = (targetStr === "" || targetStr === "me" || targetStr === "undefined") ? Session.getActiveUser().getEmail() : cmd.target;
            const parts = cmd.value.split('|');
            GmailApp.sendEmail(target, parts[0] || "AI 자동 메일", parts[1] || "내용 없음");
        },
        "2": (cmd) => {
            const limit = parseInt(cmd.value, 10) || 3;
            const threads = GmailApp.search(cmd.target, 0, limit);
            return threads.map(t => {
                const msg = t.getMessages()[0];
                return `[제목]: ${msg.getSubject()}\n[보낸사람]: ${msg.getFrom()}\n[내용]: ${msg.getPlainBody().substring(0, 300)}...`;
            }).join('\n---\n') || "해당하는 메일이 없습니다.";
        }
    }
};

/**
 * ============================================================================
 * [공용 실행 및 유틸 함수]
 * ============================================================================
 */
function processAgentResponse(jsonResponse, actionMap, sheet = null) {
    try {
        const data = typeof jsonResponse === 'string' ? JSON.parse(jsonResponse) : jsonResponse;
        const cmds = data.commands || data;
        if (!Array.isArray(cmds)) throw new Error("유효한 명령어 배열이 없습니다.");

        let results = [];
        cmds.forEach(cmd => {
            const op = String(cmd.opcode);
            if (actionMap[op]) {
                let res;
                if (sheet) {
                    const tgt = parseLtRange(sheet, cmd.target);
                    const val = parseLtRange(sheet, cmd.value) || "";
                    res = actionMap[op](sheet, tgt, val);
                } else {
                    const cal = (actionMap === WorkspaceActions.Calendar) ? (cmd.target === "PRIMARY" ? CalendarApp.getDefaultCalendar() : CalendarApp.getCalendarById(cmd.target)) : null;
                    res = actionMap[op](cmd, cal);
                }

                if (res) results.push(typeof res === 'object' ? JSON.stringify(res) : String(res));
            } else {
                throw new Error(`알 수 없는 Opcode: ${op}`);
            }
        });

        if (actionMap === WorkspaceActions.Drive) {
            try { syncDriveDbToJson("AI폴더"); } catch(e) { Logger.log("DB 싱크 실패: " + e.message); }
        }
        return results;
    } catch (e) {
        Logger.log("Agent 응답 처리 실패: " + e.message);
        throw e; // 시트 제어 실패 시 상위로 던지기 위함
    }
}

function parseLtRange(sheet, rangeStr) {
    if (!rangeStr || typeof rangeStr !== 'string' || !rangeStr.includes("lt")) return rangeStr;

    let lastRow = Utils.setlastRow(sheet);

    // 만약 B열이 완전히 비어있다면 헤더(2행)를 기준으로 설정
    if (lastRow < 2) lastRow = 2;
    // -----------------------------------------------

    return rangeStr.replace(/lt([A-Z]+)(\d+)/g, (match, col, offset) => {
        const offsetNum = parseInt(offset, 10);

        // ltB0 : 실제 데이터 다음 행 (B열 마지막행 + 1)
        // ltB1 : 실제 데이터가 있는 마지막 행
        const targetRow = (offsetNum === 0) ? lastRow + 1 : lastRow + (offsetNum - 1);

        return col + targetRow;
    });
}

function handleOpcode7(sheet, targetRange, inputValue) {
    if (!inputValue) return;
    const dataArray = inputValue.split('|').map(row => row.split(','));
    const startCell = sheet.getRange(targetRange);
    sheet.getRange(startCell.getRow(), startCell.getColumn(), dataArray.length, dataArray[0].length).setValues(dataArray);
}

function handleOpcode8(sheet, targetRange) {
    const rowIdx = parseInt(targetRange.replace(/[^0-9]/g, ''), 10);
    if (rowIdx > 2) sheet.deleteRow(rowIdx);
}

function handleOpcode9(sheet, targetRange) {
    const rowIdx = parseInt(targetRange.replace(/[^0-9]/g, ''), 10);
    if (rowIdx > 2) sheet.insertRowBefore(rowIdx);
}

function getDriveItemById(id) {
    try { return DriveApp.getFileById(id); } catch (e) { return DriveApp.getFolderById(id); }
}