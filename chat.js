// #region [선언부] - 프롬프트 원본 100% 유지
/**
 * [이벤트 엔트리포인트] 시트 편집 발생 시 실행되는 트리거 함수
 * * 최적화 전략:
 * 1. Guard Clause: 불필요한 연산 전 조기 리턴 (Lazy Evaluation)
 * 2. Batch Operations: 루프 내부 API 호출을 제거하고 메모리 내 배열 연산으로 대체 (O(1))
 * 3. Context Object: 검증 서비스에 필요한 데이터를 객체로 캡슐화하여 전달
 * * @param {Object} e - 구글 시트 편집 이벤트 객체
 */
function onEdit(e) {
    // 1. [Safety Check] 이벤트 객체가 없으면 즉시 종료
    if (!e) return;

    const range = e.range;
    const sheet = range.getSheet();

    // 2. [Feature Toggle] 보조 기능 활성화 여부 체크 (D1 셀)
    // 가장 저렴한 비용(단일 셀 조회)으로 하위의 무거운 로직 실행 여부를 결정
    if (!sheet.getRange(1, 4).getValue()) return;

    // 3. [Late Initialization] 시트 이름 확인 (API 호출 비용 발생)
    const sheetName = sheet.getName();

    // 4. [Routing] 관리자 시트인 경우 별도 컨트롤러로 위임
    if (sheetName === CONFIG.SHEET.ADMIN) {
        AdminController.handle(sheet, range);
        return;
    }

    // 5. [Pre-processing] 검증에 필요한 설정 정보 로드 (Cache 활용)
    const config = CacheManager.get();
    if (!config) return;

    // 6. [Boundary Check] 데이터가 존재하는지 확인
    const maxCol = sheet.getLastColumn();
    if (maxCol < 1) return;

    // 7. [Batch Fetching] 루프 외부에서 필요한 데이터를 한꺼번에 로드 (중요)
    const values = range.getValues();       // 수정된 영역의 실제 값들
    const startRow = range.getRow();        // 수정 시작 행 번호
    const startCol = range.getColumn();     // 수정 시작 열 번호

    // [Optimization] 수정된 행들의 전체 컬럼 데이터를 한 번의 API 호출로 가져옴 (N+1 Select 방지)
    const rowDataBatch = sheet.getRange(startRow, 1, values.length, maxCol).getValues();
    // [Optimization] 헤더 정보를 한 번에 로드하여 컬럼명 매핑에 사용
    const headers = sheet.getRange(2, 1, 1, maxCol).getValues()[0];

    // 8. [Data Processing] 수정된 모든 셀을 순회하며 검증 로직 실행
    for (let i = 0; i < values.length; i++) {
        const currentRow = startRow + i;

        // 헤더 영역(1~2행) 수정은 무시
        if (currentRow < 3) continue;

        /** * @constant {Array} activeRowFullData
         * 현재 행의 전체 데이터를 API 호출 없이 미리 로드된 배치 배열에서 추출
         */
        const activeRowFullData = rowDataBatch[i];

        for (let j = 0; j < values[i].length; j++) {
            const currentCol = startCol + j;
            const newValue = values[i][j];
            const colName = headers[currentCol - 1]; // 1-based index를 0-based array index로 변환

            // 9. [Validation Filter] 검증이 불필요한 경우 Skip
            // - 컬럼명이 없거나, 값이 비어있거나, 이전 값과 동일한 경우
            if (!colName || newValue === "" || newValue === null || (e.oldValue !== undefined && newValue === e.oldValue)) continue;

            /**
             * @constant {Object} context
             * 검증 서비스에 전달할 데이터 DTO (Data Transfer Object)
             */
            const context = {
                sheet: sheet,               // [Object] 현재 조작 중인 Google Sheet 객체
                sheetName: sheetName,       // [String] 시트 식별을 위한 시트명 (CONFIG 조회용)
                row: currentRow,            // [Number] 현재 수정 중인 셀의 절대 행 번호 (1-based)
                column: currentCol,         // [Number] 현재 수정 중인 셀의 절대 열 번호 (1-based)
                newValue: newValue,         // [Any] 사용자가 새로 입력한 값
                oldValue: e.oldValue !== undefined ? e.oldValue : "", // [Any] 수정 전의 값. 단일 셀 수정이 아닐 경우 빈 문자열로 정규화 (Null-Safety 처리)
                colName: colName,           // [String] 현재 수정 중인 컬럼의 헤더명 (검증 규칙 매핑의 키)
                headers: headers,           // [Array] 시트의 전체 헤더 목록 (인덱스 역추적용)
                config: config,             // [Object] CacheManager에서 로드된 유효성 검사 설정 데이터
                activeRowFullData: activeRowFullData // [Array] 현재 수정 중인 '행(Row)' 전체의 데이터 셋
            };

            // 10. [Business Logic] 검증 및 후속 처리 위임
            ValidationService.validate(context);
        }
    }
}

/**
 * [관리자 컨트롤러] 관리자 전용 시트의 동작을 제어
 * * 설계 포인트:
 * 1. UI 기반 트리거: 특정 셀(F1)의 변경을 감지하여 시스템 프로세스(동기화 등) 실행
 * 2. 상태 자동 복구: 작업 완료 후 트리거 셀을 초기 상태(false)로 자동 복구 (finally 블록)
 * 3. 캐시 갱신 전략: 관리자 설정 변경 시 즉시 Cache를 새로고침하여 데이터 정합성 유지
 */
const AdminController = {
    /**
     * 관리자 시트 편집 이벤트 핸들러
     * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 현재 시트 객체
     * @param {GoogleAppsScript.Spreadsheet.Range} range - 편집된 범위 객체
     */
    handle: function(sheet, range) {
        try {
            const row = range.getRow();
            const col = range.getColumn();

            if(row ===1){
                switch(col){
                    case 5: if(!getSecret("CACHE_DATA"))
                        CacheManager.refresh();break;
                    case 6:
                        // [Trigger Logic] F1 셀(1행 6열)이 수정되었을 때 시트 생성/동기화 프로세스 진입
                        if (row === 1 && col === 6) {
                            CacheManager.refresh();
                            Logger.log("[Admin] 시트 생성 프로세스 진입");
                            // 외부 모듈인 SheetManager의 존재 여부 확인 후 동기화 실행
                            if (typeof SheetManager.syncSheets === "function") SheetManager.syncSheets();
                        };
                        break;
                }
            }
            // [Cache Management] 어떤 수정이든 관리자 시트에서 발생하면 캐시를 리프레시
            //

            CacheManager.refresh();
        } catch (err) {
            Logger.log("[Admin Error] 관리 테이블 처리 실패: " + err);
        } finally {
            // [State Rollback] 트리거 셀(F1)을 다시 false로 변경하여 버튼 클릭 효과 구현
            // 자바의 리소스 해제(Cleanup) 로직과 유사한 흐름
            if (range.getRow() === 1 && range.getColumn() === 6) {
                sheet.getRange(1, 6).setValue(false);
            }
        }
    }
};

/**
 * [유효성 검증 서비스] 비즈니스 로직에 기반한 데이터 정합성 검사
 * * 설계 포인트:
 * 1. DTO 활용: onEdit에서 생성한 context 객체를 통해 필요한 모든 데이터를 주입받음
 * 2. Zero-Read: 검증 과정 중 추가적인 sheet.getRange().getValue() 호출을 원천 차단 (배치 최적화)
 * 3. Dependency Check: 현재 행의 다른 컬럼 데이터와 연계된 검증 수행 (KeyDependency)
 */
const ValidationService = {
    /**
     * 단일 셀에 대한 데이터 검증 및 후속 조치
     * @param {Object} context - 검증에 필요한 모든 데이터가 포함된 객체
     */
    validate: function(context) {
        // 1. Context 객체 구조 분해 (자바의 필드 추출과 유사)
        const { sheetName, colName, newValue, oldValue, sheet, row, column, config,activeRowFullData} = context;

        // 2. [Config Lookup] 시트명+컬럼명 조합키를 활용해 해당 셀의 검증 규칙 획득
        const configKey = `${sheetName}_${colName}`;
        const targetConfig = config.mainData[configKey];

        const co = Object.keys(config.mainData).filter((item)=>
            item.includes(sheetName)
        )
        Logger.log(co)
        // 검증 규칙이 정의되지 않은 컬럼은 Pass
        if (!targetConfig || targetConfig.type==="undefined") return;

        const type = targetConfig.type;
        const isFunctionValue = targetConfig.isFunctionValue;

        if(isFunctionValue.length > 0){return}

        let isValid = false;

        // 3. [Pattern Matching] 타입별 정규표현식 패턴 획득 및 테스트
        const pattern = this.getRegexPattern(type, config);
        Logger.log(pattern)
        if (pattern && pattern !== "boolean") {
            isValid = new RegExp(pattern).test(String(newValue));
        } else if (pattern === "boolean") {
            // 체크박스 형태의 데이터는 별도의 로직으로 유효성 관리
            isValid = false;
        }

        // 4. [Integrity Check] 상위 의존 컬럼들이 모두 채워져 있는지 최종 확인 (AND 연산)
        isValid = isValid && this.checkKeyDependency(context, targetConfig.columKey);

        // 5. [Error Handling] 검증 실패 시 데이터 복구(Rollback) 또는 UI 강제 교정
        if (!isValid) {
            Logger.log(`[Validation Reject] Cell:[${row},${column}], 입력값:${newValue}, 요구타입:${type}`);

            if (type === "boolean") {
                // 체크박스 유실 방지: 강제로 체크박스 인스턴스 재생성
                sheet.getRange(row, column).insertCheckboxes();
            } else {
                // 부적절한 값 입력 시: 이전 값(oldValue)으로 즉시 롤백 (Undo 작업)
                sheet.getRange(row, column).setValue(oldValue);
            }
        }else{
            InputFunction.validate(context)
        }
    },

    /**
     * 타입별 정규식 패턴 탐색 (CONFIG 상수 -> config 동적 데이터 순)
     */
    getRegexPattern: function(type, config) {
        if (CONFIG.REGEX[type]) {
            return CONFIG.REGEX[type];
        }
        return config.regData[type] ? config.regData[type].regex : null;
    },

    /**
     * 행 내의 특정 컬럼들에 데이터가 존재하는지 확인 (의존성 검사)
     * 최적화: API 호출 없이 미리 로드된 activeRowFullData 배열 내에서 인덱스 기반 검색
     * @param {Object} context - 실행 컨텍스트
     * @param {string} columKey - 콤마로 구분된 의존 컬럼명 리스트
     */
    checkKeyDependency: function(context, columKey) {
        // 의존성 규칙이 없으면 통과
        if (!columKey || columKey.trim() === "") return true;

        const keyList = columKey.split(',').map(item => item.trim());
        const { headers, activeRowFullData } = context;

        // 모든 의존 컬럼이 null/빈값이 아닌지 확인 (자바의 Stream.allMatch() 대응)
        return keyList.every(keyColName => {
            const colIndex = headers.indexOf(keyColName);

            if (colIndex !== -1) {
                // 로컬 배열 인덱싱을 통해 O(1) 속도로 값 확인
                const val = activeRowFullData[colIndex];
                return val !== "" && val !== null && val !== undefined;
            }
            return false; // 매핑되는 헤더가 없으면 유효하지 않은 데이터로 간주
        });
    }
};

const InputFunction = {
    validate: function (context) {
        const { sheetName, sheet, row, column, config } = context;

        // 1. 현재 시트명(sheetName) 또는 "BASE"로 시작하는 설정들만 추출
        const settings = Object.keys(config.mainData)
            .filter(key => key.startsWith(sheetName + "_") || key.startsWith("BASE_"))
            .map(key => config.mainData[key])
            .sort((a, b) => Number(a.columNum) - Number(b.columNum));

        Logger.log(`[확인] 시트(${sheetName}) + BASE 설정 합계: ${settings.length}개`);

        settings.forEach((setting) => {
            const targetCol = Number(setting.columNum) + 1;
            const val = setting.isFunctionValue;

            // 가드 클로저
            if (!val || val === "" || targetCol === column || isNaN(targetCol) || targetCol < 1) return;

            const isOnlyRow3 = (setting.isFunction === true || setting.isFunction === "true");
            let proceed = false;

            if (isOnlyRow3) {
                if (row === 3) proceed = true;
            } else {
                proceed = true;
            }

            if (proceed) {
                try {
                    // (수정) 자바스크립트 문자열 결합은 &가 아닌 + 나 템플릿 리터럴(`)을 사용합니다.
                    Logger.log(`targetCol : ${targetCol}`);

                    sheet.getRange(row, targetCol).setValue(val);
                    Logger.log(`[입력성공] 열:${targetCol}, 값:${val} (Row:${row})`);
                } catch (e) {
                    Logger.log(`[에러] ${targetCol}열 입력 중 오류: ${e.message}`);
                }
            }
        });
    }
};