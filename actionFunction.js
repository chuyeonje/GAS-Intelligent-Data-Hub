/**
 * ============================================================================
 * 데이터베이스(DB) 조작 서비스 (AlaSQL 연동용 백엔드)
 * ============================================================================
 */
const DatabaseService = {

    /**
     * 전체 DB 데이터 가져오기 (Read) - Zero Config
     * 데이터가 없는 시트도 INSERT를 위해 빈 테이블 정보(Marker)를 생성하여 전달
     */
    fetchInitialData: function() {
        let allData = [];
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // 1. 이미 mode 1에서 "관리 시트가 제외된 진짜 명단"만 가져옵니다.
        const dbSheetNames = SheetManager.getAllSheetNames(true, 1);

        Logger.log("동기화 대상 시트: " + dbSheetNames.join(", "));

        // 2. 굳이 필터링할 필요 없이, 이 명단만 가지고 바로 순회하면 됩니다.
        dbSheetNames.forEach(function(sheetName) {
            const sheet = ss.getSheetByName(sheetName);
            if (!sheet) return; // 만약의 상황 대비

            const dataCount = Number(sheet.getRange("B1").getValue());

            if (dataCount > 0) {
                const lastCol = SheetManager.getLastColNum(sheetName);
                if (lastCol > 0) {
                    const data = sheet.getRange(3, 1, dataCount, lastCol).getDisplayValues();
                    Logger.log(data)
                    Logger.log(dataCount)
                    Logger.log(lastCol)
                    allData = allData.concat(data);
                }
            } else {
                Logger.log(sheetName)
                // 데이터가 0개여도 INSERT용 마커는 보내줘야 하니까요!
                allData.push([sheetName + "/EMPTY_MARKER"]);
            }
        });
        Logger.log(allData)
        Logger.log(`[DatabaseService] 총 ${allData.length}행 로드 완료.`);
        return allData;
    },

    /**
     * 물리적 시트 데이터 조작 (Create, Update, Delete)
     * HTML에서 JSON.parse된 payload를 직접 전달받아 실행합니다.
     */
    handleMutation: function(payload) {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(payload.targetSheet);
        if (!sheet) return { success: false, error: "시트를 찾을 수 없음: " + payload.targetSheet };

        // [디버깅] 넘어온 데이터 원본 확인
        Logger.log("Payload ID: " + payload.id);
        Logger.log("Payload Action: " + payload.action);
        Logger.log("RowData: " + JSON.stringify(payload.rowData));

        const dataCount = Number(sheet.getRange("B1").getValue()) || 0;

        try {
            if (payload.action === "CREATE") {
                const insertRow = dataCount + 3; // 데이터 3행 시작 + 기존 개수 = 다음 빈 행
                // ID를 제외한 데이터 추출 시 인덱스 주의
                const dataWithoutId = payload.rowData.slice(1);
                sheet.getRange(insertRow, 2, 1, dataWithoutId.length).setValues([dataWithoutId]);
                return { success: true, message: "시트 생성 완료" };
            }

            // [중요] ID 검색 로직 강화
            // 데이터가 0개여도 수식 때문에 B1이 0이 아닐 수 있으므로 실제 범위를 넉넉히 잡음
            const lastRow = sheet.getLastRow();
            if (lastRow < 3) throw new Error("시트에 데이터가 없습니다.");

            const idValues = sheet.getRange(3, 1, lastRow - 2, 1).getValues().flat();
            // trim()을 추가하여 공백 문제 해결
            const targetIndex = idValues.findIndex(id => id.toString().trim() === payload.id.toString().trim());

            if (targetIndex === -1) {
                throw new Error(`ID를 찾을 수 없음: [${payload.id}]. 시트 내 ID 목록: ${idValues.slice(0,3)}...`);
            }

            const actualRow = targetIndex + 3;

            if (payload.action === "UPDATE") {
                if (!payload.rowData) throw new Error("업데이트할 데이터가 없습니다.");

                // [핵심] payload.rowData의 구조가 시트 A, B, C, D... 와 일치하는지 확인
                // 만약 rowData가 [ID, 이름, 카테고리, 수량] 이라면 slice(1)이 맞음
                const dataToUpdate = payload.rowData.slice(1);

                // 2번째 열(B열)부터 데이터 업데이트
                sheet.getRange(actualRow, 2, 1, dataToUpdate.length).setValues([dataToUpdate]);

                return { success: true, message: "시트 수정 완료" };
            }

            // DELETE 로직 동일...

        } catch (e) {
            Logger.log("Mutation Error: " + e.message);
            return { success: false, error: e.message };
        }
    }
};

/**
 * ============================================================================
 * 구글 드라이브 동기화 서비스
 * ============================================================================
 */
const DriveSyncService = {
    syncAiFolderToSheet: function() {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let sheet = ss.getSheetByName(SYS_CONFIG.SHEET.DRIVE_DB);

        if (!sheet) sheet = ss.insertSheet(SYS_CONFIG.SHEET.DRIVE_DB);
        sheet.clear();
        sheet.appendRow(["ID", "이름", "고유ID", "형식", "최종수정일", "부모폴더ID", "바로가기URL"]);

        const folders = DriveApp.getFoldersByName(SYS_CONFIG.FOLDER.AI);
        if (!folders.hasNext()) {
            Logger.log("[Error] AI 폴더를 찾을 수 없습니다.");
            return;
        }

        const aiFolder = folders.next();
        const parentId = aiFolder.getId();
        let dataRows = [];
        let index = 1;

        const timeZone = Session.getScriptTimeZone();

        const subFolders = aiFolder.getFolders();
        while (subFolders.hasNext()) {
            const folder = subFolders.next();
            dataRows.push([
                index++, folder.getName(), folder.getId(), "FOLDER",
                Utilities.formatDate(folder.getLastUpdated(), timeZone, "yyyy-MM-dd HH:mm:ss"),
                parentId, folder.getUrl()
            ]);
        }

        const files = aiFolder.getFiles();
        while (files.hasNext()) {
            const file = files.next();
            let type = file.getName().split('.').pop().toUpperCase();
            const mime = file.getMimeType();

            if (mime.includes('spreadsheet')) type = "SHEET";
            else if (mime.includes('document')) type = "DOC";
            else if (mime.includes('presentation')) type = "SLIDE";
            else if (type === file.getName()) type = "FILE";

            dataRows.push([
                index++, file.getName(), file.getId(), type,
                Utilities.formatDate(file.getLastUpdated(), timeZone, "yyyy-MM-dd HH:mm:ss"),
                parentId, file.getUrl()
            ]);
        }

        if (dataRows.length > 0) {
            sheet.getRange(2, 1, dataRows.length, dataRows[0].length).setValues(dataRows);
        }
        Logger.log(`[DriveSync] ${aiFolder.getName()} 폴더 스캔 완료 (${dataRows.length} 항목)`);
    }
};

/**
 * ============================================================================
 * 차트 처리 서비스
 * ============================================================================
 */
const ChartService = {
    getSelectedData : function(chartName, chartType) {
        const range = SpreadsheetApp.getActiveRange();
        if (!range) return null;

        const values = range.getValues();
        const dataObj = {};

        // 최대 5개 변수까지만 처리 (사용자 요청 사항)
        const limit = Math.min(values.length, 5);

        for (let i = 0; i < limit; i++) {
            const label = values[i][0] || `항목${i+1}`;
            const value = values[i][1] || 0;
            dataObj[label] = value;
        }

        // 최종 리턴 구조: { "차트이름": [{데이터}, "차트형태"] }
        const result = {};
        result[chartName] = [dataObj, chartType];

        return JSON.stringify(result);
    }
}

/**
 * ============================================================================
 * 오디오 처리 서비스 (Chunking & Base64) - 최종 최적화 버전
 * ============================================================================
 */
const AudioService = {
    folderIds: null,

    /**
     * 폴더 ID 초기화 및 폴더 자동 생성
     */
    init: function() {
        if (!this.folderIds) {
            this.folderIds = SYS_CONFIG.setFolderId();
        }
    },

    /**
     * 원본 폴더의 MP3 파일을 1.5MB 단위의 텍스트 조각으로 변환하여 저장
     */
    convertAllMp3ToChunks: function() {
        this.init();

        const sourceFolder = DriveApp.getFolderById(this.folderIds.MUSIC_SRC);
        const targetFolder = DriveApp.getFolderById(this.folderIds.MUSIC_TARGET);

        const existingMusicList = this.getFileList();
        const files = sourceFolder.getFilesByType("audio/mpeg");

        let processedTitles = [];

        while (files.hasNext()) {
            const file = files.next();
            const fileName = file.getName().replace(/\.[^/.]+$/, "");

            // 이미 변환된 파일은 건너뛰기
            if (existingMusicList.includes(fileName)) {
                Logger.log(`[Skip] 이미 변환됨: ${fileName}`);
                processedTitles.push(fileName);
                continue;
            }

            // GAS 메모리 제한(약 50MB)을 고려하여 40MB 초과 파일은 경고 후 처리 시도
            if (file.getSize() > 40 * 1024 * 1024) {
                Logger.log(`[Warning] 대용량 파일(${Math.round(file.getSize()/1024/1024)}MB): ${fileName}`);
            }

            try {
                const bytes = file.getBlob().getBytes();
                const quantity = this.createMusicChunks(fileName, bytes, targetFolder);
                processedTitles.push(fileName);
                Logger.log(`[Success] ${fileName} (${quantity} 조각 생성 완료)`);
            } catch (e) {
                Logger.log(`[Error] ${fileName} 변환 실패: ${e.message}`);
            }
        }

        // 최종 목록 업데이트 및 파일 저장
        const finalTitleArray = [...new Set(processedTitles)].filter(t => t.trim() !== "").sort();
        const oldTitleFiles = targetFolder.getFilesByName(SYS_CONFIG.FILE.MUSIC_LIST_TXT);
        while (oldTitleFiles.hasNext()) {
            oldTitleFiles.next().setTrashed(true);
        }
        targetFolder.createFile(SYS_CONFIG.FILE.MUSIC_LIST_TXT, finalTitleArray.join(','), MimeType.PLAIN_TEXT);
    },

    /**
     * 실제 Chunk 생성 로직
     */
    createMusicChunks: function(fileName, bytes, targetFolder) {
        const chunkSize = 1572864; // 1.5MB (Base64 변환 시 약 2MB)
        let chunkIndex = 0;

        for (let i = 0; i < bytes.length; i += chunkSize) {
            const end = Math.min(i + chunkSize, bytes.length);
            const chunkBytes = bytes.slice(i, end);
            const base64Chunk = Utilities.base64Encode(chunkBytes);
            const chunkFileName = `${fileName}_part${chunkIndex}.txt`;

            targetFolder.createFile(chunkFileName, base64Chunk, MimeType.PLAIN_TEXT);
            chunkIndex++;
        }
        return chunkIndex;
    },

    /**
     * 저장된 전체 곡 목록 가져오기
     */
    getFileList: function() {
        this.init();
        const folder = DriveApp.getFolderById(this.folderIds.MUSIC_TARGET);
        const files = folder.getFilesByName(SYS_CONFIG.FILE.MUSIC_LIST_TXT);

        if (!files.hasNext()) return [];

        const content = files.next().getBlob().getDataAsString();
        return content ? content.split(",").map(item => item.trim()).filter(item => item !== "") : [];
    },

    /**
     * 특정 곡의 총 조각 개수 확인 (HTML에서 루프 돌리기 위함)
     */
    getChunkCount: function(songTitle) {
        this.init();
        const folder = DriveApp.getFolderById(this.folderIds.MUSIC_TARGET);
        // 파일명 검색 쿼리: 해당 곡의 파트 파일들만 필터링
        const searchQuery = `title contains '${songTitle}_part' and trashed = false`;
        const files = folder.searchFiles(searchQuery);

        let count = 0;
        while (files.hasNext()) {
            files.next();
            count++;
        }
        return count;
    },

    /**
     * 특정 순번(index)의 조각 데이터 하나만 가져오기 (메모리 효율적)
     */
    getAudioChunkByIndex: function(songTitle, index) {
        this.init();
        const folder = DriveApp.getFolderById(this.folderIds.MUSIC_TARGET);
        const fileName = `${songTitle}_part${index}.txt`;
        const files = folder.getFilesByName(fileName);

        if (files.hasNext()) {
            return files.next().getBlob().getDataAsString();
        }
        return null;
    },

    /**
     * 모든 조각을 한꺼번에 가져오기
     */
    getAudioChunks: function(songTitle) {
        this.init();
        const folder = DriveApp.getFolderById(this.folderIds.MUSIC_TARGET);
        const searchQuery = `title contains '${songTitle}_part' and trashed = false`;
        const files = folder.searchFiles(searchQuery);

        let chunks = [];
        while (files.hasNext()) {
            let file = files.next();
            const match = file.getName().match(/_part(\d+)\.txt/);
            if (match) {
                chunks.push({
                    content: file.getBlob().getDataAsString(),
                    index: parseInt(match[1], 10)
                });
            }
        }
        return chunks.sort((a, b) => a.index - b.index).map(c => c.content);
    }
};

/**
 * ============================================================================
 * TextInput 보조 서비스
 * ============================================================================
 */
const TextInputService = {
    appendTextData: function(text, mode) {
        try {
            const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

            // 확장성을 고려한 Switch-Case 라우팅 구조
            switch (mode) {
                case "AI_TASK":
                    return sttAi(sheet, text);

                case "SQL_QUERY":
                    try {
                        return DatabaseService.handleMutation(JSON.parse(text));
                    } catch (e) {
                        return { success: false, error: "SQL_QUERY 모드에서는 JSON 페이로드가 필요합니다." };
                    }
                case "MEMO":
                default:
                    const activeCell = sheet.getActiveCell();
                    activeCell.setValue(text);

                    const currentRow = activeCell.getRow();
                    if (currentRow >= sheet.getMaxRows()) {
                        sheet.insertRowAfter(currentRow);
                    }

                    const nextCell = activeCell.offset(1, 0);
                    nextCell.activate();

                    return { success: true, row: currentRow, col: activeCell.getColumn() };
            }
        } catch (e) {
            Logger.log("[Error] 데이터 입력 실패: " + e.toString());
            return { success: false, error: e.toString() };
        }
    }
};
/**
 * ============================================================================
 * STT 보조 서비스
 * ============================================================================
 */
const STTService = {
    appendVoiceData: function(text, mode) {
        try {
            const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

            if (mode === "STTAI도우미" && typeof sttAi === 'function') {
                return sttAi(sheet, text);
            }

            const activeCell = sheet.getActiveCell();
            activeCell.setValue(text);

            const currentRow = activeCell.getRow();
            if (currentRow >= sheet.getMaxRows()) {
                sheet.insertRowAfter(currentRow); // 행이 모자라면 자동으로 1줄 추가
            }

            const nextCell = activeCell.offset(1, 0);
            nextCell.activate();

            return { success: true, row: currentRow, col: activeCell.getColumn() };
        } catch (e) {
            Logger.log("[Error] 음성 데이터 입력 실패: " + e.toString());
            return { success: false, error: e.toString() };
        }
    }
};

/**
 * ============================================================================
 * UI 컨트롤러 (이벤트 및 다이얼로그)
 * ============================================================================
 */
const UIController = {
    setupMenus: function() {
        const ui = SpreadsheetApp.getUi();
        ui.createMenu('💾 관리테이블 저장')
            .addItem('관리테이블 저장', 'executeManualCacheRefresh')
            .addToUi();

        ui.createMenu('🎵 음악 재생')
            .addItem('플레이어 열기', 'triggerOpenAudioPlayer')
            .addItem('노래준비', 'convertAllMp3ToChunks')
            .addToUi();

        ui.createMenu('🛠 TOOLS')
            .addItem('환경 설정', 'triggerOpenConfig')
            //.addItem('차트 생성', 'showDashboard')
            .addItem('HTML OCR 작동', 'triggerOpenOCR')
            .addItem('폴더 TEXT/OCR 작동', 'processFolderOCR')
            .addItem('Text데이터 입력기 작동', 'triggerOpenTextInput')
            //.addItem('STT데이터입력기 작동', 'triggerOpenSTTInput')
            //.addItem('STTAI도우미 작동', 'triggerOpenSTTAi')
            .addToUi();
    },

    openDialog: function(fileName, title, width, height, mode = null) {
        const html = mode ? HtmlService.createTemplateFromFile(fileName) : HtmlService.createHtmlOutputFromFile(fileName);
        if (mode) html.mode = mode;

        const htmlOutput = (mode ? html.evaluate() : html).setWidth(width).setHeight(height);
        SpreadsheetApp.getUi().showModelessDialog(htmlOutput, title);
    }
};

/**
 * ============================================================================
 * [5] 글로벌 트리거 함수 (UI 노출 API)
 * ============================================================================
 */
function onOpen() { UIController.setupMenus(); }
function syncAiFolderToSheet() { DriveSyncService.syncAiFolderToSheet(); }
function convertAllMp3ToChunks() { AudioService.convertAllMp3ToChunks(); }

function triggerOpenConfig() { if (forceAuth()) { UIController.openDialog('configSetting', 'GAS 에이전트 환경 설정', 450, 520); } }
function triggerOpenOCR() { if (forceAuth()) { UIController.openDialog('OCR', 'OCR입력기', 500, 450); } }
function triggerOpenAudioPlayer() { UIController.openDialog('music', ' ', 400, 300);}
function triggerOpenShowDashboard() {if (forceAuth()) { UIController.openDialog('chart', '차트', 800, 600); }}
function triggerOpenTextInput() {if (forceAuth()) { UIController.openDialog('InputController', '텍스트데이터입력기', 320, 380); }}
function triggerOpenSTTInput() {if (forceAuth()) { UIController.openDialog('STT', 'STT데이터입력기', 200, 320); }}
function triggerOpenSTTAi() { if (forceAuth()) { UIController.openDialog('STT', 'STTAI도우미', 200, 320); }}

// 클라이언트 사이드 연동 함수
function getAudioChunks(songTitle) { return AudioService.getAudioChunks(songTitle); }
function getAudioDataUri(fileId) { return AudioService.getAudioDataUri(fileId); }
//function appendVoiceData(text, mode) { return TextInputService.appendVoiceData(text, mode); }
function appendTextData(text, mode) { return TextInputService.appendTextData(text, mode); }
function getFileList() { return AudioService.getFileList(); }
function getChunkCount(songTitle) {return AudioService.getChunkCount(songTitle);}
function getAudioChunkByIndex(songTitle, index) {  return AudioService.getAudioChunkByIndex(songTitle, index);}
function fetchDatabaseData() { return DatabaseService.fetchInitialData(); }