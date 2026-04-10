function doPost(e) {
    try {
        // 1. 초기 요청 데이터 로그
        Logger.log("--- 요청 수신 ---");
        Logger.log("전체 이벤트 객체(e): " + JSON.stringify(e));

        const params = e.parameter;
        const fullText = params.text;
        const userName = params.user_name;

        Logger.log(`사용자: ${userName}, 입력값: "${fullText}"`);

        if (!fullText || fullText.trim() === "") {
            Logger.log("결과: 입력값 없음으로 종료");
            return createSlackResponse("⚠️ 사용법: `/ai [시트명] 작업내용` 순으로 입력해주세요.");
        }

        // 2. 텍스트 파싱 로그
        const splitText = fullText.trim().split(/\s+/);
        const targetSheetName = splitText[0];
        const aiCommand = splitText.slice(1).join(" ");

        Logger.log(`파싱 결과 - 시트명: [${targetSheetName}], 명령어: [${aiCommand}]`);

        if (!aiCommand) {
            Logger.log("결과: 작업 내용 미기입으로 종료");
            return createSlackResponse(`⚠️ 시트명 '${targetSheetName}' 뒤에 AI에게 시킬 작업을 적어주세요.`);
        }

        // 3. 스프레드시트 및 시트 확인 로그
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        Logger.log(`현재 스프레드시트 ID: ${ss.getId()}`);

        const sheet = ss.getSheetByName(targetSheetName);

        if (!sheet) {
            Logger.log(`결과: '${targetSheetName}' 시트를 찾지 못함 (실패)`);
            return createSlackResponse(`❌ 작업 실패: '${targetSheetName}' 시트를 찾을 수 없습니다.`);
        }

        // 4. 로직 진행 확인 로그
        Logger.log(`성공: '${targetSheetName}' 시트 확인 완료. AI 로직 진입 가능.`);

        const successMsg = `📁 '${targetSheetName}' 시트를 확인했습니다. AI 에이전트가 작업을 시작합니다...\n📝 요청 내용: ${aiCommand}`;

        sttAi(sheet,aiCommand);


        return createSlackResponse(successMsg);

    } catch (err) {
        // 에러 발생 시 로그 기록
        Logger.log("🚨 에러 로그: " + err.stack);
        return createSlackResponse("🚨 시스템 에러 발생: " + err.toString());
    }
}

/**
 * 슬랙 응답용 헬퍼 함수
 */
function createSlackResponse(text) {
    const result = {
        "response_type": "in_channel",
        "text": text
    };
    return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 슬랙 웹훅 전송 전용 창구
 * @param {string} message - 슬랙 채널에 표시할 내용
 */
function sendToSlack(message) {
    // 1. 이미 '한번에 처리'되어 저장되어 있을 URL을 가져옴
    const webhookUrl = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');

    if (!webhookUrl) {
        console.error("❌ 전송 실패: SLACK_WEBHOOK_URL이 설정되지 않았습니다.");
        return;
    }

    // 2. 슬랙 전송용 데이터 팩 (Payload)
    const payload = {
        "text": message,
        "username": "GAS 에이전트",
        "icon_emoji": "🤖"
    };

    // 3. HTTP 요청 설정
    const options = {
        "method": "post",
        "contentType": "application/json",
        "payload": JSON.stringify(payload),
        "muteHttpExceptions": true
    };

    // 4. 발송 실행
    try {
        const response = UrlFetchApp.fetch(webhookUrl, options);
        return response.getContentText();
    } catch (e) {
        console.error("⚠️ 슬랙 통신 에러: " + e.toString());
    }
}

// /**
//  * 전송 테스트용 (URL이 설정된 후 실행하세요)
//  */
// function debugSlack() {
//   sendToSlack("슬랙 전송 창구가 정상적으로 개설되었습니다. 🚀");
// }

function updateSlackManifest() {
    const appId = "A0123456789"; // 수정할 앱 ID
    const token = "xapp-1-..."; // App-level Token

    const manifestJson = {
        "display_information": {
            "name": "나의 자동화 봇",
            "description": "GAS로 설정을 바꿨습니다"
        },
        "features": {
            "bot_user": { "display_name": "GAS-Bot" }
        }
        // 여기에 필요한 모든 설정을 JSON으로 넣습니다.
    };

    const options = {
        "method": "post",
        "contentType": "application/json",
        "headers": { "Authorization": "Bearer " + token },
        "payload": JSON.stringify({
            "app_id": appId,
            "manifest": JSON.stringify(manifestJson) // 매니페스트는 문자열화해서 보냄
        })
    };

    const response = UrlFetchApp.fetch("https://slack.com/api/apps.manifest.update", options);
    Logger.log(response.getContentText());
}