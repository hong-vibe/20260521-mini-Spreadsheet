/**
 * 객체지향(OOP) 기반 미니 스프레드시트 애플리케이션 (script.js)
 * 
 * [고도화 업데이트 완료]
 * 1. 전체 선택 모듈: 좌측 상단 모서리(.corner-header) 클릭 시 9x9 전체 영역 일괄 활성화.
 * 2. 양방향 클립보드 연동 (Ctrl+C): 단일 셀, 전체 행, 전체 열, 전체 시트 범위에 따라 정밀 TSV 텍스트 클립보드 생성.
 * 3. 2차원 다중 셀 붙여넣기 (Ctrl+V): 외부 복사 텍스트(탭/개행 구분)를 해독하여 포커스 기준 좌표로 2차원 착지 정렬 기입 및 상태 동기화.
 * 4. 입문자 친화성: 바이브코딩 초심자도 각 기능의 흐름을 쉽게 이해하도록 모든 핵심 코드 라인에 한글 설명 제공.
 */

class SpreadsheetApp {
    /**
     * 애플리케이션 생성자: 초기 상태 정의 및 엘리먼트 캐싱을 담당합니다.
     */
    constructor() {
        // [데이터 상태 관리] 사용자가 입력한 셀 좌표와 값을 매핑하는 단일 플랫 객체
        this.spreadsheetData = {};

        // [복사/선택 범위 상태 관리] 현재 포커스된 단일 셀이나 드래그/선택된 범위 구조를 실시간 추적합니다.
        // 형태: null 또는 { type: 'cell'|'row'|'col'|'all', col: 'A', row: 1, cell: 'A1' }
        this.currentSelection = null;
        
        // [DOM 엘리먼트 캐싱] 자주 접근하는 화면 요소들을 캐시하여 성능을 극대화합니다.
        this.currentCellIndicator = document.getElementById('current-cell');
        this.exportButton = document.getElementById('export-btn');
        this.cellInputs = document.querySelectorAll('.cell-input');
        
        // [초기 구동] 이벤트 리스너 등록을 기점으로 애플리케이션을 구동시킵니다.
        this.init();
    }

    /**
     * 초기화 모듈: 바인딩 로직을 호출합니다.
     */
    init() {
        this.bindEvents();
    }

    /**
     * [이벤트 처리 모듈]
     * 화면 상의 버튼, 입력창, 헤더들에 각 이벤트를 연결하는 단일 책임 메서드입니다.
     */
    bindEvents() {
        // A. 개별 셀(Input)들에 대한 멀티 이벤트 연결
        this.cellInputs.forEach(input => {
            // 1. 셀 포커스 (마우스 클릭 또는 Tab 진입)
            input.addEventListener('focus', (e) => this.handleCellFocus(e));
            
            // 2. 셀 블러 (포커스 이탈)
            input.addEventListener('blur', () => this.handleCellBlur());
            
            // 3. 키보드 입력 감지 (엔터 및 방향키 탐색 처리)
            input.addEventListener('keydown', (e) => this.handleCellKeyDown(e));
            
            // 4. 데이터 값 실시간 입력 처리
            input.addEventListener('input', (e) => this.handleCellInput(e));
        });

        // B. 행/열 헤더(<th>) 클릭을 통한 전체 범위 선택 이벤트 바인딩
        const columnHeaders = document.querySelectorAll('th[data-header-col]');
        const rowHeaders = document.querySelectorAll('th[data-header-row]');

        columnHeaders.forEach(th => {
            th.addEventListener('click', (e) => this.handleColumnHeaderClick(e));
        });

        rowHeaders.forEach(th => {
            th.addEventListener('click', (e) => this.handleRowHeaderClick(e));
        });

        // C. 좌측 상단 모서리(.corner-header) 클릭 시 9x9 전체 선택 처리 바인딩
        const cornerHeader = document.querySelector('.corner-header');
        if (cornerHeader) {
            cornerHeader.addEventListener('click', () => this.handleCornerHeaderClick());
        }

        // D. 클립보드 복사(Ctrl+C) 및 붙여넣기(Ctrl+V) 이벤트 전역 바인딩
        document.addEventListener('copy', (e) => this.handleClipboardCopy(e));
        document.addEventListener('paste', (e) => this.handleClipboardPaste(e));

        // E. 엑셀 내보내기 버튼 클릭 이벤트
        this.exportButton.addEventListener('click', () => this.exportToExcel());
    }

    /* ==========================================
       [데이터 관리 (CRUD) 모듈 - 1함수=1역할]
       ========================================== */

    /**
     * 데이터 삽입/갱신: 전역 객체 상태에 셀 좌표별 문자열 값을 저장합니다.
     */
    updateCellValue(cellCoord, value) {
        this.spreadsheetData[cellCoord] = value;
        console.log('데이터 상태 갱신:', this.spreadsheetData);
    }

    /**
     * 데이터 삭제: 입력이 지워져 빈칸이 된 셀의 Key를 객체에서 영구 배제(delete)합니다.
     */
    deleteCellValue(cellCoord) {
        delete this.spreadsheetData[cellCoord];
        console.log('데이터 키 삭제 완료:', this.spreadsheetData);
    }

    /* ==========================================
       [UI / 시각 표현 레이어 모듈 - 1함수=1역할]
       ========================================== */

    /**
     * 실시간 셀 표시기(인디케이터) 텍스트 및 외곽선 스타일 업데이트
     */
    setIndicatorText(text, isActive = false) {
        this.currentCellIndicator.textContent = text;
        if (isActive) {
            this.currentCellIndicator.style.color = 'var(--primary-color)';
            this.currentCellIndicator.style.borderColor = 'var(--primary-color)';
        } else {
            this.currentCellIndicator.style.color = 'var(--text-muted)';
            this.currentCellIndicator.style.borderColor = '#e2e8f0';
        }
    }

    /**
     * 모든 행/열 헤더(<th>)의 하늘색 하이라이트 클래스(.active-header) 일괄 제거
     */
    clearAllHeaderHighlights() {
        const activeHeaders = document.querySelectorAll('.active-header');
        activeHeaders.forEach(header => {
            header.classList.remove('active-header');
        });
    }

    /**
     * 전체 선택 시 활성화되는 셀들의 투명 블루 배경색(.selected-cell) 일괄 제거
     */
    clearAllCellSelections() {
        const selectedCells = document.querySelectorAll('.selected-cell');
        selectedCells.forEach(cell => {
            cell.classList.remove('selected-cell');
        });
    }

    /**
     * 특정 셀에 부합하는 가로(행) 및 세로(열) 헤더 <th>에 하이라이트 적용
     */
    highlightCellHeaders(col, row) {
        // 이전 하이라이트 이력 초기화
        this.clearAllHeaderHighlights();

        const colHeader = document.querySelector(`th[data-header-col="${col}"]`);
        const rowHeader = document.querySelector(`th[data-header-row="${row}"]`);

        if (colHeader) colHeader.classList.add('active-header');
        if (rowHeader) rowHeader.classList.add('active-header');
    }

    /* ==========================================
       [셀 개별 이벤트 핸들러 모듈]
       ========================================== */

    /**
     * 셀 포커스 핸들러: 실시간 표시기 갱신, 헤더 하이라이팅 및 복사 범위 갱신
     */
    handleCellFocus(e) {
        // 기존의 행/열 전체 범위 선택이 있었다면 시각 효과 즉시 해제
        this.clearAllCellSelections();

        const col = e.target.dataset.col;
        const row = parseInt(e.target.dataset.row, 10);
        const cellCoord = e.target.dataset.cell;

        // UI 갱신
        this.setIndicatorText(`Cell: ${cellCoord}`, true);
        this.highlightCellHeaders(col, row);

        // [복사 범위 관리] 현재 복사 타겟을 이 셀 하나로 제한 지정
        this.currentSelection = {
            type: 'cell',
            col: col,
            row: row,
            cell: cellCoord
        };
    }

    /**
     * 셀 포커스 이탈 핸들러: 시각 효과 초기화 및 선택 상태 초기화
     */
    handleCellBlur() {
        // 하이라이트 해제 및 기본 텍스트 원복
        this.clearAllHeaderHighlights();
        this.setIndicatorText("Cell: 선택 안 됨", false);
        
        // 포커스 아웃 시 타겟 선택 해제 (단, 복사를 연이어 바로 하도록 하기 위해
        // 약간의 딜레이를 주거나 떼어내지 않고 유지하는 경우가 많지만,
        // 헤더 클릭 범위나 기타 상태들과의 일관성을 맞추기 위해 지표만 안전히 동기화)
    }

    /**
     * 실시간 텍스트 인풋 감지 핸들러: 데이터 추가/삭제 로직 처리
     */
    handleCellInput(e) {
        const cellCoord = e.target.dataset.cell;
        const value = e.target.value.trim();

        if (value !== '') {
            // 값이 채워져 있으면 데이터 업데이트 메서드 실행
            this.updateCellValue(cellCoord, value);
        } else {
            // 빈 칸으로 변경 시 데이터 완전 소멸 메서드 실행
            this.deleteCellValue(cellCoord);
        }
    }

    /* ==========================================
       [키보드 네비게이션 (엔터 & 방향키 이동) 모듈]
       ========================================== */

    /**
     * 키보드 이벤트 분기 처리: 엔터 키 및 방향키의 이동 제어
     */
    handleCellKeyDown(e) {
        const col = e.target.dataset.col;
        const row = parseInt(e.target.dataset.row, 10);
        
        switch (e.key) {
            case 'Enter':
                // 1. 엔터 키: 수직 아래 칸으로 포커스 자동 하강 이동
                e.preventDefault(); // 엔터 키 기본 개행/이동 방지
                this.moveFocus(col, row + 1);
                break;
                
            case 'ArrowUp':
                // 2. 위쪽 방향키: 위 칸 이동
                e.preventDefault();
                this.moveFocus(col, row - 1);
                break;
                
            case 'ArrowDown':
                // 3. 아래쪽 방향키: 아래 칸 이동
                e.preventDefault();
                this.moveFocus(col, row + 1);
                break;
                
            case 'ArrowLeft':
                // 4. 왼쪽 방향키: 왼쪽 열 이동 (아스키코드로 알파벳 마이너스 연산)
                e.preventDefault();
                const prevCol = String.fromCharCode(col.charCodeAt(0) - 1);
                this.moveFocus(prevCol, row);
                break;
                
            case 'ArrowRight':
                // 5. 오른쪽 방향키: 오른쪽 열 이동 (아스키코드로 알파벳 플러스 연산)
                e.preventDefault();
                const nextCol = String.fromCharCode(col.charCodeAt(0) + 1);
                this.moveFocus(nextCol, row);
                break;
        }
    }

    /**
     * 특정 좌표로 포커스를 이동하고, 입력창 내부의 내용을 전체 선택(select)해 즉시 덮어쓰기 타이핑을 지원합니다.
     */
    moveFocus(col, row) {
        // [경계 조건 방어 코드] 행 범위(1~9)와 열 범위(A~I)를 철저히 검증합니다.
        if (row < 1 || row > 9) return; 
        if (col < 'A' || col > 'I') return; 

        // 이동할 대상 셀의 Input 태그 탐색
        const targetInput = document.querySelector(`input[data-cell="${col}${row}"]`);
        
        if (targetInput) {
            targetInput.focus();
            // 포커스 진입 시 텍스트 전체 선택 효과를 부여해 바로 타이핑이 가능하게 유도합니다.
            targetInput.select(); 
        }
    }

    /* ==========================================
       [행/열/모서리 헤더 클릭 시 범위 선택 모듈]
       ========================================== */

    /**
     * 열(세로) 헤더 클릭 핸들러: 전체 열 범위 선택 및 복사 범위 지정
     */
    handleColumnHeaderClick(e) {
        // 1. 기존의 시각적 하이라이트들 일괄 정리
        this.clearAllHeaderHighlights();
        this.clearAllCellSelections();

        // 2. 클릭된 세로 열 문자 파악 (예: "C")
        const col = e.target.dataset.headerCol;
        
        // 3. 해당 열 헤더에 활성화 하이라이트 입히기
        e.target.classList.add('active-header');

        // 4. 동일한 열 문자 속성을 가진 모든 <input> 셀을 탐색하여 소프트 블루 배경 주입
        const targetCells = document.querySelectorAll(`input[data-col="${col}"]`);
        targetCells.forEach(cell => {
            cell.classList.add('selected-cell');
        });

        // 5. 좌표 표시기를 해당 열 전체 선택 상태로 표현
        this.setIndicatorText(`Cell: ${col}열 전체 선택됨`, true);

        // [복사 범위 관리] 해당 열 전체를 복사 타겟으로 지정
        this.currentSelection = {
            type: 'col',
            col: col
        };
    }

    /**
     * 행(가로) 헤더 클릭 핸들러: 전체 행 범위 선택 및 복사 범위 지정
     */
    handleRowHeaderClick(e) {
        // 1. 기존 하이라이트 초기화
        this.clearAllHeaderHighlights();
        this.clearAllCellSelections();

        // 2. 클릭된 가로 행 번호 파악 (예: "5")
        const row = parseInt(e.target.dataset.headerRow, 10);

        // 3. 해당 행 헤더 활성화 하이라이트 주입
        e.target.classList.add('active-header');

        // 4. 동일한 행 번호 속성을 지닌 가로라인 <input> 셀들 일괄 선택 효과
        const targetCells = document.querySelectorAll(`input[data-row="${row}"]`);
        targetCells.forEach(cell => {
            cell.classList.add('selected-cell');
        });

        // 5. 좌표 표시기를 해당 행 전체 선택 상태로 갱신
        this.setIndicatorText(`Cell: ${row}행 전체 선택됨`, true);

        // [복사 범위 관리] 해당 행 전체를 복사 타겟으로 지정
        this.currentSelection = {
            type: 'row',
            row: row
        };
    }

    /**
     * 모서리(코너) 최상단 좌측 셀 클릭 핸들러: 9x9 전체 영역 선택 및 복사 지정
     */
    handleCornerHeaderClick() {
        // 1. 시각 스타일 초기화
        this.clearAllHeaderHighlights();
        this.clearAllCellSelections();

        // 2. 9행(1~9) x 9열(A~I)을 가리키는 모든 헤더 요소를 활성화
        const allThHeaders = document.querySelectorAll('.spreadsheet-table th:not(.corner-header)');
        allThHeaders.forEach(th => {
            th.classList.add('active-header');
        });

        // 3. 81개 전체 셀 입력창을 선택 하이라이트 배경색으로 변경
        this.cellInputs.forEach(cell => {
            cell.classList.add('selected-cell');
        });

        // 4. 인디케이터 표시 갱신
        this.setIndicatorText("Cell: 시트 전체 선택됨", true);

        // [복사 범위 관리] 시트의 가로/세로 전체를 복사 대상으로 지정
        this.currentSelection = {
            type: 'all'
        };
    }

    /* ==========================================
       [양방향 클립보드 복사 / 붙여넣기 전용 모듈]
       ========================================== */

    /**
     * 복사(Ctrl+C) 이벤트 가로채기 핸들러: 
     * 선택된 영역 구조(단일 셀, 전체 행, 전체 열, 전체 시트)에 어울리는 정밀한 TSV 텍스트를 만들어 클립보드에 주입합니다.
     */
    handleClipboardCopy(e) {
        // 복사 타겟 정보가 아예 없다면 중단합니다.
        if (!this.currentSelection) return;

        let copyText = '';
        const cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'];

        // 복사된 범위 유형에 맞춰 데이터 가공 분기 진행 (TSV 포맷화)
        switch (this.currentSelection.type) {
            case 'cell':
                // A. 단일 셀 복사: 해당 셀의 데이터 값만 반환
                const coord = this.currentSelection.cell;
                copyText = this.spreadsheetData[coord] || '';
                break;

            case 'col':
                // B. 세로 열 전체 복사: 1행부터 9행까지 순회하며 수직 형태(개행 구분)로 구성
                const colLetter = this.currentSelection.col;
                const colValues = [];
                for (let r = 1; r <= 9; r++) {
                    const cCoord = `${colLetter}${r}`;
                    colValues.push(this.spreadsheetData[cCoord] || '');
                }
                copyText = colValues.join('\r\n');
                break;

            case 'row':
                // C. 가로 행 전체 복사: A열부터 I열까지 순회하며 수평 형태(탭 구분)로 구성
                const rowNum = this.currentSelection.row;
                const rowValues = [];
                cols.forEach(c => {
                    const rCoord = `${c}${rowNum}`;
                    rowValues.push(this.spreadsheetData[rCoord] || '');
                });
                copyText = rowValues.join('\t');
                break;

            case 'all':
                // D. 시트 전체 복사: 9x9 2차원 데이터를 탭(\t)과 개행(\r\n)으로 바인딩
                const gridRows = [];
                for (let r = 1; r <= 9; r++) {
                    const rowCells = [];
                    cols.forEach(c => {
                        const cellCoord = `${c}${r}`;
                        rowCells.push(this.spreadsheetData[cellCoord] || '');
                    });
                    gridRows.push(rowCells.join('\t'));
                }
                copyText = gridRows.join('\r\n');
                break;
        }

        // 획득한 데이터를 운영체제 클립보드 객체에 텍스트 타입으로 덮어씁니다.
        e.clipboardData.setData('text/plain', copyText);
        
        // 브라우저 자체 드래그 범위 복사 등의 기본 억제 동작 차단
        e.preventDefault();
        
        console.log("클립보드에 스프레드시트 호환 복사 완료:", copyText);
    }

    /**
     * 붙여넣기(Ctrl+V) 이벤트 가로채기 핸들러:
     * 외부 스프레드시트(구글 시트, 엑셀)에서 복사해 온 TSV 구조를 파싱해 활성 셀 원점(0,0)을 기준으로 2차원 분할 착지시킵니다.
     */
    handleClipboardPaste(e) {
        // 현재 활성화된 포커스 셀이 없거나 단일 셀이 아니라면 붙여넣기를 수행할 원점이 모호하므로 중단합니다.
        if (!this.currentSelection || this.currentSelection.type !== 'cell') return;

        // 클립보드로부터 원본 텍스트 데이터 획득
        const clipboardData = e.clipboardData || window.clipboardData;
        const pastedText = clipboardData.getData('text');
        
        if (!pastedText) return;

        // 붙여넣기를 전개할 시작 원점 셀 정보 획득
        const startCol = this.currentSelection.col;
        const startRow = this.currentSelection.row;
        
        const cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'];
        const startColIndex = cols.indexOf(startCol);

        // [TSV 데이터 해독기]
        // 1. 줄바꿈(\r\n 또는 \n)을 근거로 가로 행 데이터들을 분할합니다.
        let lines = pastedText.split(/\r?\n/);
        
        // 엑셀 등에서 뒤따르는 최하단 불필요한 개행 공백 행이 있으면 필터링해 줍니다.
        if (lines.length > 0 && lines[lines.length - 1] === '') {
            lines.pop();
        }

        // 2. 가로 행 내부 데이터들을 탭(\t) 문자로 구분해 2차원 배열 배열(Matrix)을 빌드합니다.
        const dataMatrix = lines.map(line => line.split('\t'));

        // 3. 브라우저 본연의 디폴트 붙여넣기 억제 (기본 작동 시 한 인풋 상자 내에 모든 글자가 뭉쳐 기입됩니다)
        e.preventDefault();

        // 4. 2차원 횡렬 순회 기입 및 상태 동기화 작동
        for (let i = 0; i < dataMatrix.length; i++) {
            // 타겟 행 행 계산
            const targetRow = startRow + i;
            
            // [경계 조건 제어] 9행 범위를 초과하는 수직 데이터는 소실 무시 처리
            if (targetRow > 9) break;

            for (let j = 0; j < dataMatrix[i].length; j++) {
                // 타겟 열 열 인덱스 계산
                const targetColIndex = startColIndex + j;
                
                // [경계 조건 제어] I열(인덱스 8)을 초과하는 가로 데이터는 무시 처리
                if (targetColIndex > 8) break;

                const targetColLetter = cols[targetColIndex];
                const targetCoord = `${targetColLetter}${targetRow}`;
                const cellValue = dataMatrix[i][j].trim();

                // 화면 입력 필드 DOM 탐색 및 값 동기화
                const cellInput = document.querySelector(`input[data-cell="${targetCoord}"]`);
                if (cellInput) {
                    cellInput.value = cellValue;
                }

                // 데이터 비즈니스 로직 동시 갱신 (빈값 여부에 따라 delete 또는 update 처리)
                if (cellValue !== '') {
                    this.updateCellValue(targetCoord, cellValue);
                } else {
                    this.deleteCellValue(targetCoord);
                }
            }
        }

        console.log("외부 클립보드 다중 셀 2차원 붙여넣기 착지 완료.");
    }

    /* ==========================================
       [SheetJS 연동 엑셀 익스포트 모듈]
       ========================================== */

    /**
     * 전역 상태 spreadsheetData를 구글 스프레드시트 규격 A1:I9 좌표계로 정밀 변환하여 파일 다운로드를 트리거합니다.
     */
    exportToExcel() {
        try {
            // A. 새로운 워크북 패키지 구축
            const wb = XLSX.utils.book_new();

            // B. 데이터와 셀 서식을 매핑할 임시 시트 컨테이너 객체
            const ws = {};

            // [구글 스프레드시트 100% 호환 장치]
            // 데이터 유무와 상관없이 물리적으로 가로 A~I, 세로 1~9 범위 크기를 강제 고정하여 업로드 시 왜곡 현상을 완벽 방제합니다.
            ws['!ref'] = "A1:I9";

            // C. 누적된 데이터를 순회하며 좌표 셀마다 텍스트 타입('s')으로 정보 저장
            for (const [cellCoord, val] of Object.entries(this.spreadsheetData)) {
                ws[cellCoord] = {
                    v: val,
                    t: 's' // 문자열 형태 지정
                };
            }

            // D. 작성된 시트를 워크북 객체에 Sheet1이라는 명칭으로 주입
            XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

            // E. 브라우저 실제 로컬 하드 드라이브로 파일(spreadsheet.xlsx) 다운로드 강제 유도
            XLSX.writeFile(wb, "spreadsheet.xlsx");

            console.log("엑셀 다운로드 처리 완료: spreadsheet.xlsx가 로컬로 저장되었습니다.");
        } catch (error) {
            console.error("SheetJS 파일 내보내기 장애 발생:", error);
            alert("엑셀 파일 변환 도중 예상치 못한 오류가 발생했습니다.");
        }
    }
}

// [애플리케이션 즉시 실행] DOM 로드가 끝나면 바로 클래스 객체를 생성하여 인스턴스화합니다.
document.addEventListener('DOMContentLoaded', () => {
    window.spreadsheetApp = new SpreadsheetApp();
});
