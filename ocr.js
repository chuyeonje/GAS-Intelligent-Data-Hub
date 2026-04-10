/**
 * OCR 처리기 팩토리
 */
const OCRProcessorFactory = {
    create: function(type) {
        switch(type) {
            case 'DRIVE_API':
                return new DriveOCRProcessor();
            // case 'VISION_API': return new VisionOCRProcessor(); // 확장 가능
            default:
                throw new Error("지원하지 않는 OCR 타입입니다.");
        }
    }
};

/**
 * 저장소 팩토리
 */
const DataStoreFactory = {
    create: function(type, config) {
        switch(type) {
            case 'GOOGLE_SHEET':
                return new SheetDataStore(config.sheetName);
            // case 'LOG_ONLY': return new LogDataStore(); // 확장 가능
            default:
                throw new Error("지원하지 않는 저장소 타입입니다.");
        }
    }
};

/**
 * Drive API를 이용한 OCR 구현체
 */
function DriveOCRProcessor() {
    this.process = function(blob, language = 'ko') {
        let tempFileId = null;
        try {
            const resource = { title: "temp_" + new Date().getTime(), mimeType: blob.getContentType() };
            const tempFile = Drive.Files.insert(resource, blob, { ocr: true, ocrLanguage: language });
            tempFileId = tempFile.id;

            const doc = DocumentApp.openById(tempFileId);
            const text = doc.getBody().getText();
            return text.split('\n').filter(line => line.trim() !== "");
        } finally {
            if (tempFileId) Drive.Files.remove(tempFileId);
        }
    };
}

/**
 * 구글 시트 저장 구현체
 */
function SheetDataStore(sheetName) {
    this.save = function(lines) {
        if (!lines || lines.length === 0) return;
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
        if (!sheet) throw new Error(`'${sheetName}' 시트를 찾을 수 없습니다.`);

        const lastRow = Utils.setlastRow(sheet);
        const startRow = lastRow < 3 ? 3 : lastRow + 1;
        const values = lines.map(line => [line]);
        sheet.getRange(startRow, 2, values.length, 1).setValues(values);
    };
}

/**
 * HTML UI에서 직접 호출하는 함수 (브릿지 역할)
 */
function ProcessUiImage(base64Data, contentType) {
    try {
        // 1. Base64 데이터를 Blob으로 변환 (이것이 자바의 Input stream 처리와 비슷합니다)
        const decodedData = Utilities.base64Decode(base64Data);
        const blob = Utilities.newBlob(decodedData, contentType, "ui_upload_image");

        // 2. 팩토리를 통해 OCR 처리기 생성
        const ocrProcessor = OCRProcessorFactory.create('DRIVE_API');

        // 3. 처리 실행 및 결과 반환
        const extractedLines = ocrProcessor.process(blob);

        // 배열을 다시 하나의 문자열로 합쳐서 UI에 전달
        return extractedLines.join('\n');

    } catch (e) {
        console.error("UI OCR 처리 중 오류: " + e.toString());
        throw new Error("OCR 분석 실패: " + e.message);
    }
}