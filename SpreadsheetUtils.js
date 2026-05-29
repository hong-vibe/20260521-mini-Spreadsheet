/**
 * SpreadsheetUtils.js
 * 
 * [스프레드시트 특수 도구함]
 * 1. 2차원 TSV 클립보드 파싱 및 다중 셀 번역
 * 2. SheetJS 기반 Excel(.xlsx) 내보내기 및 A1 포맷 가로세로 고정
 * 3. A-Z 알파벳과 0-based 열 인덱스 상호 변환기
 * 
 * 이 파일은 웹화면 DOM(HTML 상자)을 만지지 않는 순수한 수학적 계산 및 데이터 가공 부품입니다.
 */

export const SpreadsheetUtils = {
    /**
     * 0-based 열 인덱스를 알파벳 문자로 변환합니다.
     * 예: 0 -> 'A', 25 -> 'Z', 26 -> 'AA'
     * 
     * @param {number} index - 0부터 시작하는 숫자 열 인덱스
     * @returns {string} 알파벳 열 헤더 이름
     */
    getColLetter(index) {
        let temp = index;
        let letter = '';
        while (temp >= 0) {
            letter = String.fromCharCode((temp % 26) + 65) + letter;
            temp = Math.floor(temp / 26) - 1;
        }
        return letter;
    },

    /**
     * 알파벳 열 문자열을 0-based 열 인덱스로 변환합니다.
     * 예: 'A' -> 0, 'Z' -> 25, 'AA' -> 26
     * 
     * @param {string} letter - 알파벳 열 헤더 이름
     * @returns {number} 0부터 시작하는 숫자 열 인덱스
     */
    getColIndex(letter) {
        let index = 0;
        for (let i = 0; i < letter.length; i++) {
            index = index * 26 + (letter.charCodeAt(i) - 64);
        }
        return index - 1;
    },

    /**
     * 탭(TSV)으로 분리된 2차원 클립보드 텍스트를 파싱하여 행렬 배열로 반환합니다.
     * 외부 스프레드시트 복사 데이터를 유연하게 해독합니다.
     * 
     * @param {string} tsvText - 클립보드에서 읽어온 탭/개행 구분 문자열
     * @returns {Array<Array<string>>} 2차원 텍스트 배열
     */
    parseTSV(tsvText) {
        let lines = tsvText.split(/\r?\n/);
        // 맨 뒷부분의 찌꺼기 빈 줄 제거
        if (lines.length > 0 && lines[lines.length - 1] === '') {
            lines.pop();
        }
        return lines.map(line => line.split('\t'));
    },

    /**
     * 전역 상태 spreadsheetData 데이터를 SheetJS(xlsx.full.min.js)와 연동하여 
     * 구글 스프레드시트 100% 호환 규격(A1:I9 고정 또는 maxCols:maxRows 고정)으로 엑셀 다운로드를 실행합니다.
     * 
     * @param {Object} spreadsheetData - 셀 주소별 문자열 데이터 상태 객체
     * @param {number} maxCols - 현재 활성화된 최대 가로 열의 개수
     * @param {number} maxRows - 현재 활성화된 최대 세로 행의 개수
     */
    exportToExcel(spreadsheetData, maxCols, maxRows) {
        // XLSX 전역 라이브러리가 로드되어 있는지 검증
        if (typeof XLSX === 'undefined') {
            throw new Error("SheetJS 라이브러리(XLSX)가 브라우저에 적재되지 않았습니다.");
        }

        // A. 새로운 워크북 패키지 구축
        const wb = XLSX.utils.book_new();

        // B. 데이터를 매핑할 시트 컨테이너 객체
        const ws = {};

        // C. 물리적으로 A1부터 마지막 열/행까지 영역을 강제 고정하여 업로드 시 왜곡 현상 방지
        const lastColLetter = this.getColLetter(maxCols - 1);
        ws['!ref'] = `A1:${lastColLetter}${maxRows}`;

        // D. 누적된 데이터를 순회하며 좌표 셀마다 텍스트 타입('s')으로 정보 저장
        for (const [cellCoord, val] of Object.entries(spreadsheetData)) {
            // 현재 활성 차원(maxCols, maxRows) 범위를 벗어난 이전 데이터 찌꺼기는 엑셀 출력에서 배제
            const col = cellCoord.match(/[A-Z]+/)[0];
            const row = parseInt(cellCoord.match(/[0-9]+/)[0], 10);
            const colIdx = this.getColIndex(col);
            
            if (colIdx >= 0 && colIdx < maxCols && row >= 1 && row <= maxRows) {
                ws[cellCoord] = {
                    v: val,
                    t: 's' // 문자열 형태 지정
                };
            }
        }

        // E. 작성된 시트를 워크북 객체에 Sheet1 명칭으로 주입
        XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

        // F. 로컬 디스크로 즉시 파일(spreadsheet.xlsx) 다운로드 강제 트리거
        XLSX.writeFile(wb, "spreadsheet.xlsx");
    }
};
