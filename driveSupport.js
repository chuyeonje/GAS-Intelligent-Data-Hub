
/**
 * 특정 폴더(AI 폴더) 및 그 '모든 하위 폴더'의 목록을 스캔하여 시트에 기록합니다.
 * 컬럼 구조: [ID(순번), 이름, 고유ID, 형식, 최종수정일, 부모폴더ID, 바로가기URL]
 */
function syncAiFolderToSheet() {
    const sheetName = SYS_CONFIG.FOLDER.AI;
    const TARGET_FOLDER_NAME = sheetName;
    const SHEET_NAME = sheetName;

    const ss = SpreadsheetApp.getActive();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
        sheet = ss.insertSheet(SHEET_NAME);
    }

    // 1. 기존 데이터 삭제
    const lastRow = Utils.setlastRow(sheet);
    if (lastRow >= 3) {
        sheet.getRange(3, 1, lastRow - 2, 7).clearContent();
    }

    // 2. 대상 폴더(AI 폴더) 찾기
    const folders = DriveApp.getFoldersByName(TARGET_FOLDER_NAME);
    if (!folders.hasNext()) {
        Logger.log("AI 폴더를 찾을 수 없습니다.");
        return;
    }
    const rootAiFolder = folders.next();

    let dataRows = [];
    let index = 1;

    // 3. 내부 재귀 탐색 함수 정의 (하위 폴더까지 파고드는 로직)
    function scanFolder(currentFolder) {
        const currentFolderId = currentFolder.getId();

        // 폴더 스캔
        const subFolders = currentFolder.getFolders();
        while (subFolders.hasNext()) {
            const folder = subFolders.next();
            dataRows.push([
                index++,
                folder.getName(),
                folder.getId(),
                "FOLDER",
                Utilities.formatDate(folder.getLastUpdated(), "GMT+9", "yyyy-MM-dd HH:mm:ss"),
                currentFolderId, // 직계 부모 폴더 ID
                folder.getUrl()
            ]);

            // [핵심] 찾은 하위 폴더를 대상으로 다시 탐색 시작 (재귀 호출)
            scanFolder(folder);
        }

        // 파일 스캔
        const files = currentFolder.getFiles();
        while (files.hasNext()) {
            const file = files.next();
            let type = file.getName().split('.').pop().toUpperCase();

            const mime = file.getMimeType();
            if (mime.includes('spreadsheet')) type = "SHEET";
            else if (mime.includes('document')) type = "DOC";
            else if (mime.includes('presentation')) type = "SLIDE";
            else if (type === file.getName()) type = "FILE"; // 확장자 없는 경우

            dataRows.push([
                index++,
                file.getName(),
                file.getId(),
                type,
                Utilities.formatDate(file.getLastUpdated(), "GMT+9", "yyyy-MM-dd HH:mm:ss"),
                currentFolderId, // 직계 부모 폴더 ID
                file.getUrl()
            ]);
        }
    }

    // 4. 루트 AI 폴더부터 탐색 시작
    scanFolder(rootAiFolder);

    // 5. 시트 3행부터 한 번에 데이터 입력
    if (dataRows.length > 0) {
        sheet.getRange(3, 1, dataRows.length, dataRows[0].length).setValues(dataRows);
    }

    Logger.log(`${rootAiFolder.getName()} 폴더 전체 스캔 완료: 총 ${dataRows.length}개 항목 (하위 포함)`);
}


/**
 * AI 폴더 및 하위 폴더들과 시트 DB를 비교하여 실시간 상태 동기화
 */
function syncDriveDbRealtime() {
    const sheetName = SYS_CONFIG.FOLDER.AI;
    const TARGET_FOLDER_NAME = sheetName;
    const SHEET_NAME = sheetName;
    const START_ROW = 3;

    const ss = SpreadsheetApp.getActive();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return;

    // 1. 시트 데이터 로드 (고유ID 기준 매핑)
    const lastRow = Utils.setlastRow(sheet);
    const sheetMap = new Map();

    if (lastRow >= START_ROW) {
        const idValues = sheet.getRange(START_ROW, 3, lastRow - START_ROW + 1, 1).getValues();
        idValues.forEach((row, index) => {
            if (row[0]) sheetMap.set(row[0], START_ROW + index);
        });
    }

    // 2. 드라이브 전체 데이터 스캔
    const folders = DriveApp.getFoldersByName(TARGET_FOLDER_NAME);
    if (!folders.hasNext()) return;
    const rootAiFolder = folders.next();

    const currentDriveIds = new Set();
    const newItems = [];

    // 내부 재귀 탐색 함수 (실시간 비교용)
    function scanFolderRealtime(currentFolder) {
        const currentFolderId = currentFolder.getId();

        // 폴더 스캔
        const subFolders = currentFolder.getFolders();
        while (subFolders.hasNext()) {
            const folder = subFolders.next();
            const folderId = folder.getId();
            currentDriveIds.add(folderId);

            // 시트에 없는 ID라면 신규 리스트에 추가
            if (!sheetMap.has(folderId)) {
                newItems.push([
                    "", // 빈값 (ID는 함수/수식에 의해 자동 입력된다고 가정)
                    folder.getName(),
                    folderId,
                    "FOLDER",
                    Utilities.formatDate(folder.getLastUpdated(), "GMT+9", "yyyy-MM-dd HH:mm:ss"),
                    currentFolderId, // 직계 부모
                    folder.getUrl()
                ]);
            }

            // 하위 폴더 재귀 스캔
            scanFolderRealtime(folder);
        }

        // 파일 스캔
        const files = currentFolder.getFiles();
        while (files.hasNext()) {
            const file = files.next();
            const fileId = file.getId();
            currentDriveIds.add(fileId);

            // 시트에 없는 ID라면 신규 리스트에 추가
            if (!sheetMap.has(fileId)) {
                let type = file.getName().split('.').pop().toUpperCase();
                const mime = file.getMimeType();
                if (mime.includes('spreadsheet')) type = "SHEET";
                else if (mime.includes('document')) type = "DOC";
                else if (mime.includes('presentation')) type = "SLIDE";
                else if (type === file.getName()) type = "FILE";

                newItems.push([
                    "", // 빈값
                    file.getName(),
                    fileId,
                    type,
                    Utilities.formatDate(file.getLastUpdated(), "GMT+9", "yyyy-MM-dd HH:mm:ss"),
                    currentFolderId, // 직계 부모
                    file.getUrl()
                ]);
            }
        }
    }

    // 루트 폴더부터 재귀 탐색 시작
    scanFolderRealtime(rootAiFolder);

    // 3. [삭제 실행] 시트에는 있지만 드라이브에는 없는 항목 제거
    const rowsToDelete = [];
    for (let [id, rowNum] of sheetMap) {
        if (!currentDriveIds.has(id)) {
            rowsToDelete.push(rowNum);
        }
    }

    rowsToDelete.sort((a, b) => b - a).forEach(rowNum => {
        sheet.deleteRow(rowNum);
    });

    // 4. [추가 실행] 신규 항목을 시트 하단에 일괄 추가
    if (newItems.length > 0) {
        sheet.getRange(sheet.Utils.setlastRow(sheet) + 1, 1, newItems.length, 7).setValues(newItems);
    }

    Logger.log(`[하위폴더 동기화 완료] 신규 추가: ${newItems.length}건 / 삭제: ${rowsToDelete.length}건`);
}

/**
 * AI 폴더 내의 Drive_file_list.json 파일과 드라이브 상태를 실시간 동기화
 */
function syncDriveDbToJson() {
    const TARGET_FOLDER_NAME = SYS_CONFIG.FOLDER.AI;
    const DB_FILE_NAME = "Drive_file_list.json"; // 텍스트 기반 JSON 파일

    // 1. AI 폴더 찾기
    const folders = DriveApp.getFoldersByName(TARGET_FOLDER_NAME);
    if (!folders.hasNext()) {
        Logger.log(`'${TARGET_FOLDER_NAME}' 폴더를 찾을 수 없습니다.`);
        return;
    }
    const aiFolder = folders.next();

    // 2. 기존 JSON 파일 찾기 및 생성
    let dbFile;
    const files = aiFolder.getFilesByName(DB_FILE_NAME);

    if (files.hasNext()) {
        dbFile = files.next();
    } else {
        // 파일이 없으면 빈 배열 "[]"을 담은 일반 텍스트 파일로 새로 생성
        dbFile = aiFolder.createFile(DB_FILE_NAME, "[]", MimeType.PLAIN_TEXT);
        Logger.log("새로운 JSON DB 파일을 생성했습니다.");
    }

    const dbFileId = dbFile.getId();
    const currentDriveData = []; // 새롭게 기록할 전체 데이터 배열

    // 3. 드라이브 전체 데이터 스캔 (재귀 함수)
    function scanFolderRealtime(currentFolder) {
        const currentFolderId = currentFolder.getId();

        // 폴더 스캔
        const subFolders = currentFolder.getFolders();
        while (subFolders.hasNext()) {
            const folder = subFolders.next();
            const folderId = folder.getId();

            // JSON 객체 형태로 데이터 푸시
            currentDriveData.push({
                name: folder.getName(),
                id: folderId,
                type: "FOLDER",
                lastUpdated: Utilities.formatDate(folder.getLastUpdated(), "GMT+9", "yyyy-MM-dd HH:mm:ss"),
                parentId: currentFolderId,
                url: folder.getUrl()
            });

            scanFolderRealtime(folder); // 하위 폴더 재귀 호출
        }

        // 파일 스캔
        const fileIterator = currentFolder.getFiles();
        while (fileIterator.hasNext()) {
            const file = fileIterator.next();
            const fileId = file.getId();

            // DB 파일(자신)은 목록 기록에서 제외
            if (fileId === dbFileId) continue;

            // 확장자 및 MIME 타입 체크
            let type = file.getName().split('.').pop().toUpperCase();
            const mime = file.getMimeType();
            if (mime.includes('spreadsheet')) type = "SHEET";
            else if (mime.includes('document')) type = "DOC";
            else if (mime.includes('presentation')) type = "SLIDE";
            else if (type === file.getName()) type = "FILE";

            // JSON 객체 형태로 데이터 푸시
            currentDriveData.push({
                name: file.getName(),
                id: fileId,
                type: type,
                lastUpdated: Utilities.formatDate(file.getLastUpdated(), "GMT+9", "yyyy-MM-dd HH:mm:ss"),
                parentId: currentFolderId,
                url: file.getUrl()
            });
        }
    }

    // AI 루트 폴더부터 탐색 시작
    scanFolderRealtime(aiFolder);

    // 4. 추출한 데이터를 JSON 문자열로 변환하여 텍스트 파일에 덮어쓰기
    // JSON.stringify의 세 번째 인자 '2'는 사람이 읽기 편하게 들여쓰기(Indent)를 적용합니다.
    const jsonString = JSON.stringify(currentDriveData, null, 2);
    dbFile.setContent(jsonString);

    Logger.log(`[JSON 동기화 완료] 총 ${currentDriveData.length}개의 항목이 파일에 기록되었습니다.`);
}

/**
 * Drive_file_list.json 파일의 데이터를 읽어와
 * "값 | 값 | 값" 형태의 단일 텍스트로 반환하는 함수
 */
function getJsonDataAsText(folderName) {
    const TARGET_FOLDER_NAME = folderName;
    const DB_FILE_NAME = "Drive_file_list.json";

    // 1. AI 폴더 및 JSON 파일 찾기
    const folders = DriveApp.getFoldersByName(TARGET_FOLDER_NAME);
    if (!folders.hasNext()) return "폴더를 찾을 수 없습니다.";
    const aiFolder = folders.next();

    const files = aiFolder.getFilesByName(DB_FILE_NAME);
    if (!files.hasNext()) return "데이터 파일이 없습니다.";
    const dbFile = files.next();

    // 2. 텍스트 데이터 읽기 및 JSON 파싱
    const jsonString = dbFile.getBlob().getDataAsString();
    if (!jsonString || jsonString.trim() === "") return "데이터가 비어있습니다.";

    let dataArray;
    try {
        dataArray = JSON.parse(jsonString); // 문자열을 JavaScript 객체 배열로 변환
    } catch (e) {
        Logger.log("JSON 파싱 에러: " + e.message);
        return "데이터 포맷 오류입니다.";
    }

    if (dataArray.length === 0) return "추출할 데이터가 없습니다.";

    // 3. 데이터를 기존 시트 출력과 동일한 "A | B | C" 텍스트 형태로 변환
    // 헤더(열 제목) 생성
    const headers = ["이름", "ID", "유형", "수정일", "직계부모ID", "URL"].join(" | ");

    // 각 JSON 객체의 값을 배열로 뽑아낸 뒤 " | "로 연결
    const rows = dataArray.map(item => {
        return [
            item.name || "",
            item.id || "",
            item.type || "",
            item.lastUpdated || "",
            item.parentId || "",
            item.url || ""
        ].join(" | ");
    });

    // 헤더와 데이터 행들을 줄바꿈(\n)으로 연결하여 최종 문자열 생성
    const rawContext = headers + "\n" + rows.join("\n");

    return rawContext;
}