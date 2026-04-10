/**
 * [상수 정의]
 */
const CONFIG = {
    REGEX: {
        string: "^[a-zA-Zㄱ-ㅎ가-힣\s0-9]+$",
        number: "^-?\\d*\\.?\\d+$",
        boolean: "boolean",
        undefined: ""
    },
    SHEET: {
        ADMIN: "관리 테이블",
        SEARCH : "조회용 쿼리테이블",
        BASE : "BASE",
        BASE2 : "BASE2"
    },
    KEY : {
        "CACHE_DATA" : "관리테이블 시트 데이터의 최신화가 안되어 있습니다.",
        "FOLDER_ID": "폴더의 ID가 세팅이 안되어 있습니다.",
        "GEMINI_API_KEY": "제미나이의 API KEY가 세팅이 안되어 있습니다.",
        "SLACK_WEBHOOK_URL": "SLACK의 웹후크 URL이 세팅이 안되어 있습니다"
    }
};

const SYS_CONFIG = {
    FOLDER: {
        AI: "AI",
        MUSIC_SRC: "Musig",
        MUSIC_TARGET: "MusigList",
        ORC : "ORC"
    },
    SHEET: {
        DRIVE_DB: "DriveDB"
    },
    FILE: {
        MUSIC_LIST_TXT: "filesTitleList"
    },

    /**
     * 현재 스프레드시트가 있는 폴더 내에서 폴더를 찾고,
     * 없으면 새로 생성한 뒤 ID를 반환합니다.
     */
    setFolderId: function() {
        // 1. 현재 실행 중인 스프레드시트 파일의 부모 폴더 가져오기
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const file = DriveApp.getFileById(ss.getId());
        const parentFolders = file.getParents();

        if (!parentFolders.hasNext()) {
            throw new Error("부모 폴더를 찾을 수 없습니다. (내 드라이브 최상단일 수 있습니다)");
        }
        const parentFolder = parentFolders.next();

        const folderIds = {};
        const folderKeys = Object.keys(this.FOLDER);

        // 2. 설정된 폴더 이름들을 순회하며 확인 및 생성
        folderKeys.forEach(key => {
            const folderName = this.FOLDER[key];
            const folders = parentFolder.getFoldersByName(folderName);

            let folder;
            if (folders.hasNext()) {
                // 폴더가 이미 존재하는 경우
                folder = folders.next();
                console.log(`기존 폴더 연결: ${folderName} (${folder.getId()})`);
            } else {
                // 폴더가 없는 경우 새로 생성
                folder = parentFolder.createFolder(folderName);
                console.log(`새 폴더 생성됨: ${folderName} (${folder.getId()})`);
            }

            folderIds[key] = folder.getId();
        });

        return folderIds;
    }
};

const googleDriveObject = {
    fileId : "",
    targetFolder : "",
}

/**
 * [캐시 매니저] LockService 적용 및 예외 방어
 */
const CacheManager = {
    refresh: function() {
        const lock = LockService.getScriptLock();
        try {
            lock.waitLock(10000);

            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const adminSheet = ss.getSheetByName(CONFIG.SHEET.ADMIN);

            if (!adminSheet) {
                Logger.log("[Error] '관리 테이블' 시트를 찾을 수 없습니다.");
                return null;
            }

            const data = adminSheet.getDataRange().getValues();
            const cacheObj = {
                globalEnabled: data[0][3], // D1
                mainData: {},
                tableData: {},
                regData: {}
            };

            const defaultRegexKeys = Object.keys(CONFIG.REGEX);

            for (const row of data.slice(2)) {
                const mainKey = `${row[1]}_${row[2]}`;
                const tableKey = row[10];
                const regKey = row[15];

                cacheObj.mainData[mainKey] = {
                    columNum: row[3],
                    columKey: row[4],
                    type: row[5],
                    isFunction: row[6],
                    isFunctionValue: row[7]
                };

                if (tableKey) {
                    cacheObj.tableData[tableKey] = {
                        tableName: row[10],
                        tableBaseYn: row[11],
                        tableMaxNumber: row[12]
                    };
                }

                if (regKey) {
                    cacheObj.regData[regKey] = {
                        regex: row[15],
                        regexValue: defaultRegexKeys.includes(row[15]) ? row[16] : null,
                        type: defaultRegexKeys.includes(row[15]) ? row[17] : null
                    };
                }
            }

            PropertiesService.getScriptProperties().setProperty('CACHE_DATA', JSON.stringify(cacheObj));
            Logger.log("[CacheManager] 캐시 갱신 완료");
            return cacheObj;

        } catch (e) {
            Logger.log("[CacheManager Error] 락 획득 실패 또는 처리 오류: " + e);
            return null;
        } finally {
            lock.releaseLock();
        }
    },

    get: function() {
        const cacheStr = getSecret('CACHE_DATA');
        if (!cacheStr) {
            Logger.log("[CacheManager] 캐시가 비어있어 새로 생성합니다.");
            return this.refresh();
        }
        return JSON.parse(cacheStr);
    }
};

/**
 * [UI Controller] 커스텀 메뉴에서 호출되는 글로벌 래퍼 함수
 * 내부적으로 CacheManager 싱글톤을 호출하고 결과를 UI에 피드백합니다.
 */
function executeManualCacheRefresh() {
    const ui = SpreadsheetApp.getUi();

    try {
        // 알림: 사용자에게 작업이 시작되었음을 인지시킴 (선택 사항)
        SpreadsheetApp.getActiveSpreadsheet().toast("캐시 데이터를 동기화하는 중입니다...", "시스템 알림", 3);

        // 비즈니스 로직 호출
        const result = CacheManager.refresh();

        if (result) {
            ui.alert(
                '✅ 캐시 저장 완료',
                '관리 테이블의 최신 설정이 캐시에 성공적으로 반영되었습니다.\n이제 변경된 규칙이 즉시 적용됩니다.',
                ui.ButtonSet.OK
            );
        } else {
            ui.alert(
                '❌ 캐시 저장 실패',
                "'관리 테이블' 시트를 찾을 수 없거나 내부 락(Lock) 충돌이 발생했습니다.\n잠시 후 다시 시도해주세요.",
                ui.ButtonSet.OK
            );
        }
    } catch (e) {
        Logger.log("[UI Action Error] 수동 캐시 갱신 중 에러: " + e);
        ui.alert(
            '⚠️ 시스템 오류',
            '캐시 저장 중 예기치 못한 오류가 발생했습니다.\n상세 에러: ' + e.message,
            ui.ButtonSet.OK
        );
    }
}

/**
 * 속성 서비스에서 값을 가져오고, 없을 경우 CONFIG에 정의된 에러 메시지 출력
 */
function getSecret(keyName) {
    const props = PropertiesService.getScriptProperties();
    const val = props.getProperty(keyName);

    if (!val) {
        // CONFIG.KEY에서 keyName에 해당하는 메시지를 가져옴. 없으면 기본 메시지 출력.
        const customMsg = CONFIG.KEY[keyName] || "정의되지 않은 설정 키입니다.";
        const errorLog = `[설정 오류] ${keyName}: ${customMsg}`;

        console.error(errorLog);

        try {
            Browser.msgBox(errorLog);
        } catch (e) {
            // UI가 없는 실행 환경 대응
        }
        return null;
    }

    return val;
}

/**
 * HTML 폼에서 전달받은 데이터를 저장
 * @param {Object} data - HTML에서 보낸 설정 객체
 * @return {string} - 처리 결과 메시지
 */
function saveKeys(data) {
    const props = PropertiesService.getScriptProperties();

    try {
        // 폼에서 넘어온 데이터들을 속성 서비스에 매핑하여 저장
        const keys = Object.keys(data);

        for (let i = 0; i < keys.length; i++) {
            const item = keys[i];
            const value = data[item];

            if (value) {
                switch (item) {
                    case "geminiKey":
                        props.setProperties({ 'GEMINI_API_KEY': value });
                        break;
                    case "slackUrl":
                        props.setProperties({ 'SLACK_WEBHOOK_URL': value });
                        break;
                    default:
                        props.setProperties({ 'FOLDER_ID': value });
                }
            }
        }

        // 2. 권한 승인 상태 체크
        // REQUIRED_USER_ACTION 이라면 아직 승인이 안 된 상태임
        const authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
        const status = authInfo.getAuthorizationStatus(); // 여기서 Status를 꺼냅니다.
        if (status === ScriptApp.AuthorizationStatus.REQUIRED_USER_ACTION) {
            // 권한이 없을 때만 실행되는 블록
            try {
                // appsscript.json에 명시된 서비스들을 한 번씩 호출하여 팝업 유도
                DriveApp.getRootFolder();
                UrlFetchApp.getRequest("https://google.com");
                // 필요하다면 GmailApp이나 Calendar도 추가
            } catch (e) {
                // 여기서 구글 시스템이 사용자에게 권한 승인창을 띄웁니다.
                throw new Error("최초 1회 권한 승인이 필요합니다. 팝업창에서 [허용]을 눌러주세요.");
            }
        }

        return "✅ 모든 설정이 성공적으로 저장되었습니다!";
    } catch (e) {
        console.error("저장 실패: " + e.toString());
        return "❌ 저장 중 오류가 발생했습니다: " + e.message;
    }
}

function forceAuth() {
    // 권한 정보를 가져오려면 아래 메서드를 사용해야 합니다.
    const authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
    const status = authInfo.getAuthorizationStatus(); // 여기서 Status를 꺼냅니다.

    // 상태값이 'REQUIRED (승인 필요)'일 때만 팝업 로직을 실행합니다.
    if (status === ScriptApp.AuthorizationStatus.REQUIRED) {
        try {
            // 구글 시스템이 함수 진입 시점에 팝업을 띄우고 대기합니다.
            DriveApp.getRootFolder();
            UrlFetchApp.fetch("https://google.com");
        } catch (e) {
            console.error("권한 미승인: " + e);
            return false; // 사용자가 팝업을 닫거나 거부하면 false 반환
        }
    }

    // 이미 권한이 있거나(NOT_REQUIRED) 방금 승인을 마쳤다면 true 반환
    return true;
}


/**
 * ============================================================================
 * [3] 공통 유틸리티 (기존 nullck, createOb, deleteCells, getCells 등 통합)
 * ============================================================================
 */
const Utils = {
    // 빈 값 검증 (기존 nullck)
    isEmpty: function(value) {
        return value === null || value === undefined || String(value).trim() === "" || String(value) === 'undefined' || String(value) === 'null';
    },

    // 범위 데이터 읽기 (기존 getCells 통합)
    getRangeData: function(sheet, start, end) {
        const range = end && (end.x !== 0 || end.y !== 0) ?
            sheet.getRange(start.y, start.x, end.y - start.y + 1, end.x - start.x + 1) :
            sheet.getRange(start.y, start.x);
        return end && (end.x !== 0 || end.y !== 0) ? range.getValues() : range.getValue();
    },

    // 텍스트를 오브젝트로 변환 (기존 createOb, xy)
    parseObject: function(text) {
        if (!text) return {};
        const str = String(text);
        const colonIdx = str.indexOf(':');
        if (colonIdx === -1) return { [str]: [] };

        const key = str.substring(0, colonIdx).trim();
        const values = str.substring(colonIdx + 1).split(',').map(v => v.trim());
        return { [key]: values };
    },

    getCoordinate: function(location_x = 1, location_y = 1) {
        return { x: location_x, y: location_y };
    }
    ,
    setlastRow: function(sheet){
        return sheet.getRange("B1").getValue()+2; // B열 전체를 가져옴
    },
    /**
     * 입력된 날짜/시간 문자열에 9시간을 더해 Date 객체로 반환하는 함수
     * @param {string} dStr - "2023-10-27" 또는 "T14:30:00" 형태의 문자열
     * @returns {Date} 9시간이 더해진 Date 객체
     */
    getAdjustedDate : function (dStr){
        if (!dStr) return new Date(); // 값이 없으면 현재 시간 반환 (선택 사항)

        let dateStr = dStr.trim();

        // "T"로 시작하는 경우 오늘 날짜를 붙여줌 (기존 로직 유지)
        if (dateStr.startsWith('T')) {
            const today = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd");
            dateStr = today + dateStr;
        }

        const date = new Date(dateStr);

        // 9시간(ms 단위)을 더함: 9 * 60분 * 60초 * 1000ms
        const NINE_HOURS_MS = 9 * 60 * 60 * 1000;
        return new Date(date.getTime() + NINE_HOURS_MS);
    }
};


/**
 * ============================================================================
 * [4] 독립 호출용 커스텀 함수 (Cell 함수)
 * ============================================================================
 */

/**
 * 모든 시트 이름을 가져옵니다.
 * @customfunction
 */
function sGetSheetNames() {
    let cacheStr = getSecret('CACHE_DATA');

    if (!cacheStr) {
        // 커스텀 함수 내부에서는 Logger보다 Error를 반환하는 것이 셀 UI상 직관적임
        throw new Error("캐시 데이터가 없습니다.");
    }

    const tables = Object.keys(JSON.parse(cacheStr).tableData);

    // 1차원 배열을 2차원 배열로 변환하여 세로(Column) 방향으로 데이터가 뿌려지도록 처리
    return tables.map(tableName => [tableName]);
}

/**
 * 활성화된 시트 이름을 반환합니다.
 */
function getSheetName() {
    return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}