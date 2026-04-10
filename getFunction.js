/**
 * ============================================================================
 * [1] 비즈니스 로직: OCR 추출 및 시트 기록 (기존 setImageFile... 시리즈 통합)
 * ============================================================================
 */

/**
 * 메인 컨트롤러: 폴더 내 파일 순회 및 OCR 처리
 */
function processFolderOCR(name = "ALL", targetSheetName) {
    // 1. 초기 설정 및 시트 확인
    const resultSheet = targetSheetName ? targetSheetName : SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
    const FOLDER_ID = getSecret('FOLDER_ID');
    if(!FOLDER_ID)return
    // 데이터 버퍼 (Java의 List<String>과 동일한 역할)
    let dataBuffer = [];
    let fileCount = 0;

    try {
        const targetFolder = DriveApp.getFolderById(FOLDER_ID);
        const files = (name === "ALL") ? targetFolder.getFiles() : targetFolder.getFilesByName(name);

        if (!files.hasNext()) return Browser.msgBox("파일 없음");

        // 2. 데이터 수집 단계 (메모리 작업)
        while (files.hasNext()) {
            const file = files.next();
            const mimeType = file.getMimeType();

            try {
                let extractedLines = [];
                if (mimeType === MimeType.PLAIN_TEXT || mimeType === "text/csv") {
                    extractedLines = file.getBlob().getDataAsString().split(/\r?\n/).filter(line => line.trim());
                } else if (mimeType === MimeType.PDF || mimeType.indexOf('image/') !== -1) {
                    extractedLines = OCRService.extractText(file.getBlob());
                }

                if (extractedLines.length > 0) {
                    // 버퍼에 데이터 적재 (Memory append)
                    dataBuffer = dataBuffer.concat(extractedLines);
                    fileCount++;
                }
                Utilities.sleep(5000)
            } catch (fileErr) {
                // 개별 파일 오류는 로그만 남기고 계속 진행 (결함 허용)
                Logger.log(`[File Skip] ${file.getName()}: ${fileErr.message}`);
            }
        }

        // 3. 일괄 기록 (Batch Write / Commit)
        if (dataBuffer.length > 0) {
            SheetWriter.appendLines(resultSheet, dataBuffer);
            Browser.msgBox(`${fileCount}개 파일, 총 ${dataBuffer.length}행 일괄 처리 완료`);
        }

    } catch (globalErr) {
        // 4. 시스템 장애 시 Flush (사용자님의 핵심 아이디어)
        // 루프 도중 타임아웃이나 예상치 못한 오류 발생 시, 메모리에 남은 것만이라도 쓰고 죽음
        if (dataBuffer.length > 0) {
            Logger.log("Critical Error 발생! 현재까지 수집된 데이터를 긴급 Flush 합니다.");
            SheetWriter.appendLines(resultSheet, dataBuffer);
        }
        Logger.log(`[System Error] ${globalErr.stack}`);
        Browser.msgBox(`시스템 오류: ${globalErr.message}`);
    }
}

/**
 * 메인 컨트롤러: html에 수집한 OCR 처리
 */
function processFileUpload(formObject) {
    try {
        const fileBlob = formObject.myFile;

        // 1. OCR 처리 (제공해주신 서비스 호출)
        const lines = OCRService.extractText(fileBlob);

        if (lines.length === 0) return "추출된 텍스트가 없습니다.";

        // 2. 시트에 데이터 기록 (제공해주신 서비스 호출)
        // 시트 이름을 실제 사용 중인 시트명으로 바꾸세요 (예: 'OCR결과')
        const targetSheetName = "Sheet1";
        SheetWriter.appendLines(targetSheetName, lines);

        // 3. UI에 보여줄 결과 반환
        return {
            status: "success",
            data: lines.join('\n')
        };
    } catch (e) {
        return {
            status: "error",
            message: e.toString()
        };
    }
}

/**
 * OCR 처리 전담 서비스
 */
const OCRService = {
    extractText: function(blob) {
        const resource = { title: blob.getName(), mimeType: blob.getContentType() };
        const options = { ocr: true, ocrLanguage: 'ko' };

        let tempFileId = null;
        try {
            const tempFile = Drive.Files.insert(resource, blob, options);
            tempFileId = tempFile.id;
            const doc = DocumentApp.openById(tempFileId);
            const text = doc.getBody().getText();
            return text ? text.split('\n') : [];
        } finally {
            if (tempFileId) DriveApp.getFileById(tempFileId).setTrashed(true);
        }
    }
};

/**
 * 시트 쓰기 전담 서비스 (데이터 덮어쓰기 방지 적용)
 */
const SheetWriter = {
    appendLines: function(sheetName, lines, startCol = 2) {
        if (!lines || lines.length === 0) return;

        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
        if (!sheet) throw new Error(`시트를 찾을 수 없습니다: ${sheetName}`);

        let lastRow = Utils.setlastRow(sheet);
        let targetRow = lastRow < 2 ? 3 : lastRow + 1; // 3행부터 시작하거나 마지막 행 아래에 추가

        const values = lines.map(line => [line]);
        sheet.getRange(targetRow, startCol, lines.length, 1).setValues(values);
    }
};


/**
 * ============================================================================
 * 시트 세팅 및 구조 관리
 * ============================================================================
 */
const SheetManager = {
    syncSheets: function() {
        const cacheStr = getSecret('CACHE_DATA');
        if (!cacheStr) return;

        const config = JSON.parse(cacheStr);
        const tables = config.tableData || {};
        const targetSheetNames = Object.keys(tables);

        const currentSheetNames = this.getAllSheetNames(true);

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const base1 = ss.getSheetByName("BASE");
        const base2 = ss.getSheetByName("BASE2");

        targetSheetNames.forEach(name => {
            if (!currentSheetNames.includes(name)) {
                const base = (tables[name].tableBaseYn !== "NO") ? base1 : base2;
                if (base) {
                    const newSheet = base.copyTo(ss).setName(name).showSheet();
                    const cell = newSheet.getRange(2, 1);
                    newSheet.getRange(1, 5).setValue(name);
                    const formula = cell.getFormula();
                    if (formula) {
                        cell.setFormula(formula.replace(/&getSheetName\(\)&/g, `&"${name}"&`));
                    }
                }
            }
        });

        currentSheetNames.forEach(name => {
            if (!targetSheetNames.includes(name) &&
                name !== "관리 테이블" &&
                name !== "BASE" &&
                name !== "BASE2") {
                const ds = ss.getSheetByName(name);
                if (ds) ss.deleteSheet(ds);
            }
        });
    },

    getAllSheetNames: function(includeHidden = true, mode = 0) {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let allSheets = ss.getSheets();

        // 1. 숨김 시트 필터링 (객체 단계에서 먼저 수행)
        if (!includeHidden) {
            allSheets = allSheets.filter(s => !s.isSheetHidden());
        }

        // 2. 시트 이름만 추출
        const allSheetNames = allSheets.map(s => s.getName());

        // 3. 모드에 따른 필터링 결과 반환
        switch(mode) {
            case 1:
                // [블랙리스트 방식] CONFIG.SHEET에 있는 이름들은 가차없이 제외
                const excludeList = Object.values(CONFIG.SHEET).filter(val => val && val.trim() !== "");
                return allSheetNames.filter(name => !excludeList.includes(name));

            default:
                // 기본적으로는 필터링된 모든 시트 이름 반환
                return allSheetNames;
        }
    },
    getLastColNum: function(sheetName) {
        // 1. 문자열로 저장된 캐시 데이터를 가져옴
        const rawData = getSecret('CACHE_DATA');

        if (!rawData) return 26; // 데이터가 아예 없으면 기본값 반환

        try {
            // 2. [핵심] 문자열을 진짜 객체로 변환
            const cacheStr = JSON.parse(rawData);

            // 3. 객체 구조에서 데이터 추출 (방어 로직 포함)
            if (!cacheStr.tableData || !cacheStr.mainData) return 26;

            // tableData가 객체이므로 키(시트명)로 바로 접근
            const sheetInfo = cacheStr.tableData[sheetName];
            if (!sheetInfo) return 26;

            // mainData 객체의 키를 배열로 바꿔서 필터링
            const mainKeys = Object.keys(cacheStr.mainData);

            // 4. 컬럼 개수 계산 (sheetName 포함 + BASE 포함)
            const columnCount = mainKeys.filter(key =>
                key.includes(sheetName) || key.includes("BASE")
            ).length;

            return columnCount || 26; // 0이 나오면 안전하게 26 반환

        } catch (e) {
            console.error("getLastColNum 파싱 중 오류:", e);
            return 26;
        }
    }
};