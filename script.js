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

        // [실행 취소(Undo) 상태 관리] 이전의 데이터 모델 상태를 깊은 복사하여 순차적으로 쌓는 히스토리 스택
        this.undoStack = [];

        // [복사/선택 범위 상태 관리] 현재 포커스된 단일 셀이나 드래그/선택된 범위 구조를 실시간 추적합니다.
        // 형태: null 또는 { type: 'cell'|'row'|'col'|'all'|'range', col: 'A', row: 1, cell: 'A1', element: tdElement, startCol, startRow, endCol, endRow }
        this.currentSelection = null;
        
        // [수정 모드 상태 관리] 현재 셀 내에서 입력창이 활성화된 상태인지 추적합니다.
        this.isEditing = false;

        // [드래그 범위 선택 상태 관리] 마우스 드래그를 통한 다중 선택 영역을 정교히 추적하기 위한 플래그
        this.isDragging = false;
        this.dragStartCell = null; // { col: 'A', row: 1 }
        this.dragEndCell = null;   // { col: 'C', row: 3 }

        // [열 너비 조절 상태 관리] 피드백 반영 리사이즈 추적 멤버
        this.isResizing = false;
        this.resizingTh = null;
        this.resizeStartX = 0;
        this.resizeStartWidth = 0;
        
        // [DOM 엘리먼트 캐싱] 자주 접근하는 화면 요소들을 캐시하여 성능을 극대화합니다.
        this.currentCellIndicator = document.getElementById('current-cell');
        this.exportButton = document.getElementById('export-btn');
        this.cells = document.querySelectorAll('.spreadsheet-cell');
        
        // [초기 구동] 이벤트 리스너 등록을 기점으로 애플리케이션을 구동시킵니다.
        this.init();
    }

    /**
     * 초기화 모듈: 바인딩 로직을 호출합니다.
     */
    init() {
        this.initColumnResizers(); // <th> 리사이저 핸들 동적 배치
        this.bindEvents();
        this.loadFromLocalStorage(); // [2단계] 로컬스토리지에서 기존 데이터 안전 복구 로드
    }

    /**
     * [이벤트 처리 모듈]
     * 화면 상의 버튼, 셀, 헤더들에 각 이벤트를 연결하는 단일 책임 메서드입니다.
     */
    bindEvents() {
        // A. 개별 셀(td)들에 대한 이벤트 연결
        this.cells.forEach(cell => {
            // 1. 셀 마우스 다운 (단일 클릭 선택 대기 & 드래그 시작 & Shift 클릭 분기)
            cell.addEventListener('mousedown', (e) => this.handleCellMouseDown(e));
            
            // 2. 셀 마우스 엔터 (드래그 동작 중 범위 업데이트)
            cell.addEventListener('mouseenter', (e) => this.handleCellMouseEnter(e));
            
            // 3. 셀 더블 클릭 (즉시 수정 모드로 전환)
            cell.addEventListener('dblclick', (e) => this.handleCellDblClick(e));
        });

        // B. 전역 마우스 업 이벤트 바인딩 (드래그 피니시 감지)
        document.addEventListener('mouseup', () => this.handleDocumentMouseUp());

        // C. 전역 키보드 이벤트 바인딩 (선택 대기 상태에서의 방향키 네비게이션 및 타이핑 감지)
        document.addEventListener('keydown', (e) => this.handleDocumentKeyDown(e));

        // D. 행/열 헤더(<th>) 클릭을 통한 전체 범위 선택 이벤트 바인딩
        const columnHeaders = document.querySelectorAll('th[data-header-col]');
        const rowHeaders = document.querySelectorAll('th[data-header-row]');

        columnHeaders.forEach(th => {
            th.addEventListener('click', (e) => this.handleColumnHeaderClick(e));
        });

        rowHeaders.forEach(th => {
            th.addEventListener('click', (e) => this.handleRowHeaderClick(e));
        });

        // E. 좌측 상단 모서리(.corner-header) 클릭 시 9x9 전체 선택 처리 바인딩
        const cornerHeader = document.querySelector('.corner-header');
        if (cornerHeader) {
            cornerHeader.addEventListener('click', () => this.handleCornerHeaderClick());
        }

        // F. 원클릭 목업 주입 Sample 버튼 이벤트 바인딩 (피드백 반영)
        const sampleBtn = document.getElementById('sample-btn');
        if (sampleBtn) {
            sampleBtn.addEventListener('click', () => this.injectSampleMockupData());
        }

        // G. 클립보드 복사(Ctrl+C) 및 붙여넣기(Ctrl+V) 이벤트 전역 바인딩
        document.addEventListener('copy', (e) => this.handleClipboardCopy(e));
        document.addEventListener('paste', (e) => this.handleClipboardPaste(e));

        // H. 실행 취소(Undo) 버튼 클릭 이벤트 바인딩
        const undoBtn = document.getElementById('undo-btn');
        if (undoBtn) {
            undoBtn.addEventListener('click', () => this.undo());
        }

        // I. 엑셀 내보내기 버튼 클릭 이벤트
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
        this.saveToLocalStorage(); // [2단계] 상태 모델 변경 시 실시간 자동 저장 트리거
    }

    /**
     * 데이터 삭제: 입력이 지워져 빈칸이 된 셀의 Key를 객체에서 영구 배제(delete)합니다.
     */
    deleteCellValue(cellCoord) {
        delete this.spreadsheetData[cellCoord];
        console.log('데이터 키 삭제 완료:', this.spreadsheetData);
        this.saveToLocalStorage(); // [2단계] 상태 모델 변경 시 실시간 자동 저장 트리거
    }

    /* ==========================================
       [3단계 - 실행 취소(Undo) 엔진 모듈]
       ========================================== */

    /**
     * [3단계] 현재 spreadsheetData의 데이터 상태를 깊은 복사하여 Undo 스택에 저장합니다.
     * 메모리 누수를 방지하기 위해 최대 스택 크기를 50개로 제한합니다.
     */
    saveStateToUndoStack() {
        const snapshot = JSON.parse(JSON.stringify(this.spreadsheetData));
        this.undoStack.push(snapshot);

        // 최대 스택 크기를 50개로 제한하여 메모리 과부하 방지
        if (this.undoStack.length > 50) {
            this.undoStack.shift();
        }
        console.log('실행 취소(Undo) 스택 저장 완료. 현재 크기:', this.undoStack.length);
    }

    /**
     * [3단계] 가장 최근에 저장했던 상태로 롤백하고 화면과 로컬스토리지를 강제 동기화합니다.
     */
    undo() {
        // 편집 중인 상태라면 편집을 롤백(저장 안 함)하고 안전하게 이탈시킵니다.
        if (this.isEditing && this.currentSelection && this.currentSelection.element) {
            this.exitEditMode(this.currentSelection.element, false);
        }

        if (this.undoStack.length === 0) {
            console.log('되돌릴 이전 이력이 존재하지 않습니다.');
            return;
        }

        // 스택에서 가장 최근의 상태 복원
        const prevState = this.undoStack.pop();
        this.spreadsheetData = prevState;

        // 화면 그리드 일제히 청소
        this.cells.forEach(cell => {
            cell.textContent = '';
        });

        // 복구된 데이터 모델로 화면 그리드 재렌더링
        this.renderAllData();

        // 로컬스토리지 오토세이브 즉시 덮어쓰기 동기화
        this.saveToLocalStorage();

        // 선택된 단일 셀이 있다면 헤더 하이라이트 동기화
        if (this.currentSelection && this.currentSelection.type === 'cell') {
            const col = this.currentSelection.col;
            const row = this.currentSelection.row;
            this.highlightCellHeaders(col, row);
        } else {
            this.clearAllHeaderHighlights();
        }

        console.log('실행 취소(Undo) 완료. 남은 스택 크기:', this.undoStack.length);
    }

    /* ==========================================
       [2단계 - 로컬스토리지 자동 저장 및 복구 엔진 모듈]
       ========================================== */

    /**
     * [2단계] 현재의 전역 스프레드시트 데이터 상태를 로컬스토리지에 오토 세이브합니다.
     */
    saveToLocalStorage() {
        try {
            localStorage.setItem('pingpong_spreadsheet_data', JSON.stringify(this.spreadsheetData));
            console.log('로컬스토리지 자동 저장 성공.');
        } catch (error) {
            console.error('로컬스토리지 저장 실패:', error);
        }
    }

    /**
     * [2단계] 페이지 첫 기동 시 로컬스토리지를 검사해 기존 데이터 상태를 완벽 복구 로드합니다.
     */
    loadFromLocalStorage() {
        try {
            const saved = localStorage.getItem('pingpong_spreadsheet_data');
            if (saved) {
                this.spreadsheetData = JSON.parse(saved);
                console.log('로컬스토리지 복구 데이터 감지:', this.spreadsheetData);
                this.renderAllData(); // 화면상 모든 td에 복구된 값 렌더링
            } else {
                console.log('로컬스토리지에 저장된 이전 세션 데이터가 존재하지 않습니다.');
            }
        } catch (error) {
            console.error('로컬스토리지 복구 처리 중 장애 발생:', error);
        }
    }

    /**
     * [2단계] 복구된 spreadsheetData 상태를 루프하여 9x9 화면 그리드 td 엘리먼트에 고스란히 뿌려줍니다.
     */
    renderAllData() {
        for (const [cellCoord, val] of Object.entries(this.spreadsheetData)) {
            const cellTd = document.querySelector(`.spreadsheet-cell[data-cell="${cellCoord}"]`);
            if (cellTd) {
                cellTd.textContent = val;
            }
        }
        console.log('복원 데이터 화면 일괄 렌더링 완성.');
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
       [셀 엑셀식 입력 UX 상태 머신 및 핸들러 모듈]
       ========================================== */

    /**
     * 셀 마우스 다운 핸들러: 드래그 범위 선택의 시발점이며, Shift+클릭 조건도 처리합니다.
     */
    handleCellMouseDown(e) {
        // 이미 해당 셀을 에딧 중인 상태라면 마우스 클릭 동작 무시
        if (this.isEditing && this.currentSelection && this.currentSelection.element === e.currentTarget) {
            return;
        }

        // 다른 곳을 편집 중이었다면 안전하게 저장 후 탈출
        if (this.isEditing && this.currentSelection && this.currentSelection.element) {
            this.exitEditMode(this.currentSelection.element, true);
        }

        const td = e.currentTarget;
        const col = td.dataset.col;
        const row = parseInt(td.dataset.row, 10);

        // A. Shift 키를 누른 상태에서 클릭한 경우: 시작 원점 대비 사각형 범위를 설정함
        if (e.shiftKey) {
            e.preventDefault();
            this.handleShiftClickSelection(td);
            return;
        }

        // B. 일반적인 단일 클릭: 초록 테두리를 즉시 씌우고 드래그를 시작함
        this.selectCell(td);
        this.isDragging = true;
        this.dragStartCell = { col, row };
        this.dragEndCell = { col, row };
    }

    /**
     * 셀 마우스 엔터 핸들러: 드래그 활성화 중일 때, 마우스 궤적에 따른 영역을 하이라이트합니다.
     */
    handleCellMouseEnter(e) {
        if (!this.isDragging) return;

        const td = e.currentTarget;
        const col = td.dataset.col;
        const row = parseInt(td.dataset.row, 10);

        this.dragEndCell = { col, row };
        this.updateDragSelection();
    }

    /**
     * 전역 마우스 업 핸들러: 드래그를 정지시키고 최종 다중 사각형 선택 영역을 박제 확정합니다.
     */
    handleDocumentMouseUp() {
        if (!this.isDragging) return;
        this.isDragging = false;

        const start = this.dragStartCell;
        const end = this.dragEndCell;

        if (!start || !end) return;

        // 시작 좌표와 끝 좌표가 동일하면 단일 셀 선택 상태로 유지
        if (start.col === end.col && start.row === end.row) {
            return;
        }

        // 시작 좌표와 끝 좌표가 다르면 type: 'range' 다중 사각형 셀 영역으로 복사 상태를 지정
        this.currentSelection = {
            type: 'range',
            startCol: start.col,
            startRow: start.row,
            endCol: end.col,
            endRow: end.row,
            element: this.currentSelection.element // 최초 클릭 셀(초록 테두리 소유)을 타겟 원점으로 유지
        };

        // 실시간 인디케이터에 선택된 사각형 영역 크기를 표현
        const cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'];
        const sColIdx = cols.indexOf(start.col);
        const eColIdx = cols.indexOf(end.col);
        const colCount = Math.abs(sColIdx - eColIdx) + 1;
        const rowCount = Math.abs(start.row - end.row) + 1;
        
        this.setIndicatorText(`Cell: ${colCount}R x ${rowCount}C 범위 선택됨`, true);
    }

    /**
     * Shift + 클릭 선택 핸들러: 기존 포커싱 원점을 기준으로 사각형 영역을 계산해 일괄 선택합니다.
     */
    handleShiftClickSelection(targetTd) {
        // 기존 포커스된 원점 셀이 없을 경우 단일 선택 처리
        if (!this.currentSelection || !this.currentSelection.element) {
            this.selectCell(targetTd);
            return;
        }

        const anchorTd = this.currentSelection.element;
        const startCol = anchorTd.dataset.col;
        const startRow = parseInt(anchorTd.dataset.row, 10);
        
        const endCol = targetTd.dataset.col;
        const endRow = parseInt(targetTd.dataset.row, 10);

        this.dragStartCell = { col: startCol, row: startRow };
        this.dragEndCell = { col: endCol, row: endRow };

        // 시작점의 초록색 선택 테두리는 보존하고, 사각형 영역 업데이트 실행
        this.updateDragSelection();

        // 선택 범위 상태 박제
        this.currentSelection = {
            type: 'range',
            startCol: startCol,
            startRow: startRow,
            endCol: endCol,
            endRow: endRow,
            element: anchorTd
        };

        const cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'];
        const sColIdx = cols.indexOf(startCol);
        const eColIdx = cols.indexOf(endCol);
        const colCount = Math.abs(sColIdx - eColIdx) + 1;
        const rowCount = Math.abs(startRow - endRow) + 1;

        this.setIndicatorText(`Cell: ${colCount}R x ${rowCount}C 범위 선택됨`, true);
    }

    /**
     * dragStartCell과 dragEndCell 사이의 사각형 영역 내 모든 td에 하이라이트를 실시간 업데이트합니다.
     */
    updateDragSelection() {
        const start = this.dragStartCell;
        const end = this.dragEndCell;

        if (!start || !end) return;

        const cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'];
        const sColIdx = cols.indexOf(start.col);
        const eColIdx = cols.indexOf(end.col);

        // 최소/최대 인덱스 계산을 통해 드래그 방향에 상관없이 완벽한 사각형 획득
        const minColIdx = Math.min(sColIdx, eColIdx);
        const maxColIdx = Math.max(sColIdx, eColIdx);
        const minRow = Math.min(start.row, end.row);
        const maxRow = Math.max(start.row, end.row);

        // 기존 다중 셀 선택 스타일들만 걷어냄
        this.clearAllCellSelections();

        // 2차원 사각형을 돌며 하이라이트 클래스 주입
        for (let r = minRow; r <= maxRow; r++) {
            for (let cIdx = minColIdx; cIdx <= maxColIdx; cIdx++) {
                const colLetter = cols[cIdx];
                const targetTd = document.querySelector(`.spreadsheet-cell[data-cell="${colLetter}${r}"]`);
                if (targetTd) {
                    targetTd.classList.add('selected-cell');
                }
            }
        }
    }

    /**
     * 셀 더블 클릭 핸들러: 즉시 수정 모드로 진입합니다.
     */
    handleCellDblClick(e) {
        const td = e.currentTarget;
        this.enterEditMode(td);
    }

    /**
     * 특정 셀을 '선택 대기' 상태(엑셀 초록 테두리)로 전환하는 내부 보조 메서드입니다.
     */
    selectCell(tdElement) {
        // 기존의 모든 시각적 범위 선택 효과 해제
        this.clearAllCellSelections();
        this.clearSelectedCellClass();

        const col = tdElement.dataset.col;
        const row = parseInt(tdElement.dataset.row, 10);
        const cellCoord = tdElement.dataset.cell;

        // 초록 테두리 클래스 주입
        tdElement.classList.add('cell-selected');

        // UI 갱신
        this.setIndicatorText(`Cell: ${cellCoord}`, true);
        this.highlightCellHeaders(col, row);

        // 복사 및 키보드 입력을 위해 선택 상태 객체 갱신
        this.currentSelection = {
            type: 'cell',
            col: col,
            row: row,
            cell: cellCoord,
            element: tdElement
        };
    }

    /**
     * 기존의 초록 테두리(.cell-selected) 클래스 일괄 제거
     */
    clearSelectedCellClass() {
        const selected = document.querySelectorAll('.cell-selected');
        selected.forEach(el => {
            el.classList.remove('cell-selected');
        });
    }

    /**
     * [수정 모드 진입]: td 내에 임시 input을 생성해 에딧 모드로 변환합니다.
     * @param {HTMLElement} tdElement 에딧을 시작할 td 엘리먼트
     * @param {string|null} initialValue 선택 대기 상태에서 타이핑하여 즉시 유입된 덮어쓰기용 첫 글자 값
     */
    enterEditMode(tdElement, initialValue = null) {
        if (this.isEditing) return;
        this.isEditing = true;

        // td에 에딧 활성화 클래스 동적 주입 (style.css와 연동하여 0패딩으로 인풋창이 꽉 차게 됨)
        tdElement.classList.add('cell-editing-active');

        const cellCoord = tdElement.dataset.cell;
        
        // A. 기존에 노출되던 순수 텍스트 값 획득 (혹은 저장된 데이터 상태값)
        const originalVal = this.spreadsheetData[cellCoord] || '';

        // B. 임시 input 태그 생성 및 커스텀 스타일 클래스 바인딩
        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'cell-input-edit';

        // C. 초기 타이핑 문자가 존재하면 덮어쓰고, 없으면 기존 값을 로드함
        if (initialValue !== null) {
            input.value = initialValue;
        } else {
            input.value = originalVal;
        }

        // D. td의 텍스트 노드를 비우고 동적 인풋 삽입
        tdElement.textContent = '';
        tdElement.appendChild(input);
        
        // E. 즉시 포커스 유도
        input.focus();

        // 덮어쓰기가 아닌 더블클릭/엔터 진입일 경우 기존 내용 전체 선택 활성화
        if (initialValue === null) {
            input.select();
        }

        // F. 동적 인풋에 대한 이벤트 밀착 바인딩
        input.addEventListener('keydown', (e) => {
            const col = tdElement.dataset.col;
            const row = parseInt(tdElement.dataset.row, 10);

            if (e.key === 'Enter') {
                // 엔터 입력 시: 값 저장 및 탈출 후 아래쪽 셀로 강제 네비게이션
                e.preventDefault();
                this.exitEditMode(tdElement, true);
                this.moveFocus(col, row + 1);
            } else if (e.key === 'Escape') {
                // ESC 입력 시: 값 롤백(저장 안함) 및 수정 상태 복구 탈출
                e.preventDefault();
                this.exitEditMode(tdElement, false);
                // 다시 초록색 선택 상태로 강제 전환
                this.selectCell(tdElement);
            } else if (e.key === 'Tab') {
                // 탭 입력 시: 값 저장 후 오른쪽 셀로 네비게이션
                e.preventDefault();
                this.exitEditMode(tdElement, true);
                const nextCol = String.fromCharCode(col.charCodeAt(0) + 1);
                this.moveFocus(nextCol, row);
            }
        });

        // 포커스 아웃(마우스 다른 곳 클릭 등) 발생 시 안전하게 저장 후 소멸
        input.addEventListener('blur', () => {
            this.exitEditMode(tdElement, true);
        });
    }

    /**
     * [수정 모드 이탈]: 동적 input을 제거하고 원래의 td 텍스트 렌더링으로 롤백/저장합니다.
     */
    exitEditMode(tdElement, shouldSave) {
        if (!this.isEditing) return;

        // td에 에딧 활성화 클래스 즉시 해제 (style.css와 연동하여 td 고유 패딩이 복원됨)
        tdElement.classList.remove('cell-editing-active');

        const cellCoord = tdElement.dataset.cell;
        const input = tdElement.querySelector('.cell-input-edit');
        
        if (!input) {
            this.isEditing = false;
            return;
        }

        let finalValue = input.value.trim();

        if (shouldSave) {
            const originalVal = this.spreadsheetData[cellCoord] || '';
            // [3단계] 실제로 값이 변경되었을 때만 실행 취소 스택에 사전 백업 저장
            if (finalValue !== originalVal) {
                this.saveStateToUndoStack();
            }

            if (finalValue !== '') {
                // 데이터 갱신 및 화면 주입
                this.updateCellValue(cellCoord, finalValue);
                tdElement.textContent = finalValue;
            } else {
                // 빈 칸일 경우 상태 삭제
                this.deleteCellValue(cellCoord);
                tdElement.textContent = '';
            }
        } else {
            // 취소 모드일 경우 기존 데이터로 텍스트 환원
            tdElement.textContent = this.spreadsheetData[cellCoord] || '';
        }

        // 상태값 초기화
        this.isEditing = false;
    }

    /* ==========================================
       [키보드 네비게이션 및 전역 덮어쓰기 감지 모듈]
       ========================================== */

    /**
     * 전역 도큐먼트 키다운 이벤트 분기 처리:
     * 선택 대기 상태에서의 방향키 이동 및 타이핑 시작 시 덮어쓰기 진입 처리
     */
    /**
     * 전역 도큐먼트 키다운 이벤트 분기 처리:
     * 선택 대기 상태에서의 방향키 이동 및 타이핑 시작 시 덮어쓰기 진입 처리
     */
    handleDocumentKeyDown(e) {
        // [3단계] Ctrl + Z 단축키 감지 시 실행 취소(Undo) 실행 (단, 편집 중이 아닐 때)
        if (e.ctrlKey && (e.key === 'z' || e.key === 'Z')) {
            e.preventDefault();
            this.undo();
            return;
        }

        // 만약 셀을 편집(에딧)하고 있는 도중에는 전역 방향키 리스너 작동을 배제합니다.
        if (this.isEditing) return;

        // 선택 대기 중인 영역이 아예 없다면 중단합니다.
        if (!this.currentSelection) return;

        // A. [피드백 반영] Delete 또는 Backspace 입력 시 선택 영역 전체 일괄 삭제 처리
        if (e.key === 'Delete' || e.key === 'Backspace') {
            e.preventDefault();
            this.clearSelectionValues();
            return;
        }

        // 방향키 및 문자 입력 덮어쓰기는 단일 셀 선택 상태('cell')일 때만 허용합니다.
        if (this.currentSelection.type !== 'cell') return;

        const col = this.currentSelection.col;
        const row = this.currentSelection.row;
        const currentTd = this.currentSelection.element;

        // B. 방향키 및 주요 특수 키 분기
        switch (e.key) {
            case 'Enter':
                // 대기 상태에서 엔터를 누르면 더블클릭 효과로 편집 모드 진입
                e.preventDefault();
                this.enterEditMode(currentTd);
                return;

            case 'ArrowUp':
                e.preventDefault();
                this.moveFocus(col, row - 1);
                return;
                
            case 'ArrowDown':
                e.preventDefault();
                this.moveFocus(col, row + 1);
                return;
                
            case 'ArrowLeft':
                e.preventDefault();
                const prevCol = String.fromCharCode(col.charCodeAt(0) - 1);
                this.moveFocus(prevCol, row);
                return;
                
            case 'ArrowRight':
                e.preventDefault();
                const nextCol = String.fromCharCode(col.charCodeAt(0) + 1);
                this.moveFocus(nextCol, row);
                return;
        }

        // C. 덮어쓰기(Overwrite) 타이핑 즉시 감지 장치:
        // 단일 문자 입력이면서 메타/단축키(Ctrl, Alt 등)가 개입되지 않았을 때만 작동
        if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
            e.preventDefault();
            // 첫 글자를 품고 즉시 수정 모드로 폭풍 진입합니다!
            this.enterEditMode(currentTd, e.key);
        }
    }

    /**
     * 특정 좌표로 포커스(선택 대기)를 즉시 안전 이동시키는 제어 함수입니다.
     */
    moveFocus(col, row) {
        // [경계 조건 방어 코드] 행 범위(1~9)와 열 범위(A~I)를 철저히 검증합니다.
        if (row < 1 || row > 9) return; 
        if (col < 'A' || col > 'I') return; 

        // 이동할 대상 셀의 td 태그 탐색
        const targetTd = document.querySelector(`.spreadsheet-cell[data-cell="${col}${row}"]`);
        
        if (targetTd) {
            this.selectCell(targetTd);
        }
    }

    /**
     * [피드백 반영] 현재 선택 범위(단일 셀, 범위 드래그, 전체 행/열 등)의 셀 텍스트 및 전역 데이터를 일괄 제거합니다.
     */
    clearSelectionValues() {
        if (!this.currentSelection) return;

        const cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'];

        // [3단계] 지울 데이터가 실제로 존재하는지 사전 검사
        let hasDataToDelete = false;
        switch (this.currentSelection.type) {
            case 'cell':
                if (this.spreadsheetData[this.currentSelection.cell]) hasDataToDelete = true;
                break;
            case 'range':
                const sColIdx = cols.indexOf(this.currentSelection.startCol);
                const eColIdx = cols.indexOf(this.currentSelection.endCol);
                const minColIdx = Math.min(sColIdx, eColIdx);
                const maxColIdx = Math.max(sColIdx, eColIdx);
                const minRow = Math.min(this.currentSelection.startRow, this.currentSelection.endRow);
                const maxRow = Math.max(this.currentSelection.startRow, this.currentSelection.endRow);
                for (let r = minRow; r <= maxRow; r++) {
                    for (let cIdx = minColIdx; cIdx <= maxColIdx; cIdx++) {
                        if (this.spreadsheetData[`${cols[cIdx]}${r}`]) {
                            hasDataToDelete = true;
                            break;
                        }
                    }
                    if (hasDataToDelete) break;
                }
                break;
            case 'col':
                const colLetter = this.currentSelection.col;
                for (let r = 1; r <= 9; r++) {
                    if (this.spreadsheetData[`${colLetter}${r}`]) {
                        hasDataToDelete = true;
                        break;
                    }
                }
                break;
            case 'row':
                const rowNum = this.currentSelection.row;
                for (let cIdx = 0; cIdx < cols.length; cIdx++) {
                    if (this.spreadsheetData[`${cols[cIdx]}${rowNum}`]) {
                        hasDataToDelete = true;
                        break;
                    }
                }
                break;
            case 'all':
                if (Object.keys(this.spreadsheetData).length > 0) hasDataToDelete = true;
                break;
        }

        // 지울 데이터가 존재할 때만 실행 취소 스택에 추가
        if (hasDataToDelete) {
            this.saveStateToUndoStack();
        }

        switch (this.currentSelection.type) {
            case 'cell':
                // 1. 단일 셀 삭제
                const coord = this.currentSelection.cell;
                this.deleteCellValue(coord);
                if (this.currentSelection.element) {
                    this.currentSelection.element.textContent = '';
                }
                break;

            case 'range':
                // 2. 다중 드래그 사각형 범위 일괄 삭제
                const sColIdx = cols.indexOf(this.currentSelection.startCol);
                const eColIdx = cols.indexOf(this.currentSelection.endCol);
                const minColIdx = Math.min(sColIdx, eColIdx);
                const maxColIdx = Math.max(sColIdx, eColIdx);
                
                const minRow = Math.min(this.currentSelection.startRow, this.currentSelection.endRow);
                const maxRow = Math.max(this.currentSelection.startRow, this.currentSelection.endRow);

                for (let r = minRow; r <= maxRow; r++) {
                    for (let cIdx = minColIdx; cIdx <= maxColIdx; cIdx++) {
                        const colLetter = cols[cIdx];
                        const cellCoord = `${colLetter}${r}`;
                        this.deleteCellValue(cellCoord);
                        const td = document.querySelector(`.spreadsheet-cell[data-cell="${cellCoord}"]`);
                        if (td) td.textContent = '';
                    }
                }
                break;

            case 'col':
                // 3. 열 전체 일괄 삭제
                const colLetter = this.currentSelection.col;
                for (let r = 1; r <= 9; r++) {
                    const cellCoord = `${colLetter}${r}`;
                    this.deleteCellValue(cellCoord);
                    const td = document.querySelector(`.spreadsheet-cell[data-cell="${cellCoord}"]`);
                    if (td) td.textContent = '';
                }
                break;

            case 'row':
                // 4. 행 전체 일괄 삭제
                const rowNum = this.currentSelection.row;
                cols.forEach(c => {
                    const cellCoord = `${c}${rowNum}`;
                    this.deleteCellValue(cellCoord);
                    const td = document.querySelector(`.spreadsheet-cell[data-cell="${cellCoord}"]`);
                    if (td) td.textContent = '';
                });
                break;

            case 'all':
                // 5. 시트 전체 초기화 삭제
                for (let r = 1; r <= 9; r++) {
                    cols.forEach(c => {
                        const cellCoord = `${c}${r}`;
                        this.deleteCellValue(cellCoord);
                        const td = document.querySelector(`.spreadsheet-cell[data-cell="${cellCoord}"]`);
                        if (td) td.textContent = '';
                    });
                }
                break;
        }
        console.log("선택 범위 일괄 삭제 처리 완료.");
    }

    /* ==========================================
       [피드백 반영 - 열 가로 너비(Width) 조절 마우스 조작 모듈]
       ========================================== */

    /**
     * 각 열 헤더<th> 우측 끝단에 드래그 리사이즈 핸들 DOM을 동적 부착합니다.
     */
    initColumnResizers() {
        const columnHeaders = document.querySelectorAll('th[data-header-col]');
        columnHeaders.forEach(th => {
            const handle = document.createElement('div');
            handle.className = 'resize-handle';
            th.appendChild(handle);

            // 열 너비 변경 드래그 트리거 바인딩
            handle.addEventListener('mousedown', (e) => this.handleResizeMouseDown(e, th));
        });
    }

    /**
     * 리사이즈 드래그 시작 mousedown
     */
    handleResizeMouseDown(e, th) {
        e.preventDefault();
        e.stopPropagation(); // 헤더 자체 클릭(열 전체 선택) 이벤트 전파를 철저히 억제합니다.

        this.isResizing = true;
        this.resizingTh = th;
        this.resizeStartX = e.clientX;
        this.resizeStartWidth = th.offsetWidth;

        // document 전체에 전역 이동 및 마우스 뗌 이벤트를 결합
        this.resizeMouseMoveHandler = (moveEvt) => this.handleResizeMouseMove(moveEvt);
        this.resizeMouseUpHandler = () => this.handleResizeMouseUp();

        document.addEventListener('mousemove', this.resizeMouseMoveHandler);
        document.addEventListener('mouseup', this.resizeMouseUpHandler);
    }

    /**
     * 리사이즈 드래그 진행 mousemove
     */
    handleResizeMouseMove(e) {
        if (!this.isResizing || !this.resizingTh) return;

        const deltaX = e.clientX - this.resizeStartX;
        const newWidth = Math.max(50, this.resizeStartWidth + deltaX); // 최소 너비 50px 방어선 구축

        // 테이블 레이아웃 고유 크기 스타일 실시간 동적 적용
        this.resizingTh.style.width = `${newWidth}px`;
    }

    /**
     * 리사이즈 드래그 종료 mouseup
     */
    handleResizeMouseUp() {
        if (!this.isResizing) return;
        this.isResizing = false;
        this.resizingTh = null;

        // 등록된 전역 리스너 소멸 처리 (메모리 누수 원천 방지)
        document.removeEventListener('mousemove', this.resizeMouseMoveHandler);
        document.removeEventListener('mouseup', this.resizeMouseUpHandler);
        console.log("열 너비 리사이징 동작 완료.");
    }

    /* ==========================================
       [행/열/모서리 헤더 클릭 시 범위 선택 모듈]
       ========================================== */

    /**
     * 열(세로) 헤더 클릭 핸들러: 전체 열 범위 선택 및 복사 범위 지정
     */
    handleColumnHeaderClick(e) {
        // 만약 열 리사이징 조작 핸들을 클릭한 경우라면 범위 선택 동작을 취소합니다.
        if (e.target.classList.contains('resize-handle')) return;

        // 1. 기존의 시각적 하이라이트들 일괄 정리
        this.clearAllHeaderHighlights();
        this.clearAllCellSelections();
        this.clearSelectedCellClass();

        // 2. 클릭된 세로 열 문자 파악 (예: "C")
        const col = e.target.dataset.headerCol;
        
        // 3. 해당 열 헤더에 활성화 하이라이트 입히기
        e.target.classList.add('active-header');

        // 4. 동일한 열 문자 속성을 가진 모든 td 셀을 탐색하여 소프트 블루 배경 주입
        const targetCells = document.querySelectorAll(`.spreadsheet-cell[data-col="${col}"]`);
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
        this.clearSelectedCellClass();

        // 2. 클릭된 가로 행 번호 파악 (예: "5")
        const row = parseInt(e.target.dataset.headerRow, 10);

        // 3. 해당 행 헤더 활성화 하이라이트 주입
        e.target.classList.add('active-header');

        // 4. 동일한 행 번호 속성을 지닌 가로라인 td 셀들 일괄 선택 효과
        const targetCells = document.querySelectorAll(`.spreadsheet-cell[data-row="${row}"]`);
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
        this.clearSelectedCellClass();

        // 2. 9행(1~9) x 9열(A~I)을 가리키는 모든 헤더 요소를 활성화
        const allThHeaders = document.querySelectorAll('.spreadsheet-table th:not(.corner-header)');
        allThHeaders.forEach(th => {
            th.classList.add('active-header');
        });

        // 3. 81개 전체 td 셀을 선택 하이라이트 배경색으로 변경
        this.cells.forEach(cell => {
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

            case 'range':
                // B. 마우스 드래그 혹은 Shift+클릭 다중 사각형 셀 복사: 선택 사각형 매트릭스를 정밀 TSV 변환
                const rCols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'];
                const sColIdx = rCols.indexOf(this.currentSelection.startCol);
                const eColIdx = rCols.indexOf(this.currentSelection.endCol);
                const minColIdx = Math.min(sColIdx, eColIdx);
                const maxColIdx = Math.max(sColIdx, eColIdx);
                
                const minRow = Math.min(this.currentSelection.startRow, this.currentSelection.endRow);
                const maxRow = Math.max(this.currentSelection.startRow, this.currentSelection.endRow);
                
                const rangeRows = [];
                for (let r = minRow; r <= maxRow; r++) {
                    const rowCells = [];
                    for (let cIdx = minColIdx; cIdx <= maxColIdx; cIdx++) {
                        const colLetter = rCols[cIdx];
                        const cellCoord = `${colLetter}${r}`;
                        rowCells.push(this.spreadsheetData[cellCoord] || '');
                    }
                    rangeRows.push(rowCells.join('\t'));
                }
                copyText = rangeRows.join('\r\n');
                break;

            case 'col':
                // C. 세로 열 전체 복사: 1행부터 9행까지 순회하며 수직 형태(개행 구분)로 구성
                const colLetter = this.currentSelection.col;
                const colValues = [];
                for (let r = 1; r <= 9; r++) {
                    const cCoord = `${colLetter}${r}`;
                    colValues.push(this.spreadsheetData[cCoord] || '');
                }
                copyText = colValues.join('\r\n');
                break;

            case 'row':
                // D. 가로 행 전체 복사: A열부터 I열까지 순회하며 수평 형태(탭 구분)로 구성
                const rowNum = this.currentSelection.row;
                const rowValues = [];
                cols.forEach(c => {
                    const rCoord = `${c}${rowNum}`;
                    rowValues.push(this.spreadsheetData[rCoord] || '');
                });
                copyText = rowValues.join('\t');
                break;

            case 'all':
                // E. 시트 전체 복사: 9x9 2차원 데이터를 탭(\t)과 개행(\r\n)으로 바인딩
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

        e.preventDefault();

        // [피드백 반영] 공통 2차원 해독 모듈(importTSVData)로 파싱 처리를 완전히 이관 단일화
        this.importTSVData(pastedText, startCol, startRow);
        console.log("외부 클립보드 다중 셀 2차원 붙여넣기 착지 완료.");
    }

    /**
     * [피드백 반영] 탭(TSV)으로 구분된 데이터를 좌표에 맞춰 2차원 해독 렌더링하고 상태 모델에 기입하는 핵심 공통 모듈입니다.
     */
    importTSVData(tsvText, startCol, startRow, skipUndo = false) {
        // [3단계] 스킵 플래그가 꺼져 있을 때만(예: Ctrl+V 붙여넣기 등) 실행 취소 백업 진행
        if (!skipUndo) {
            this.saveStateToUndoStack();
        }

        const cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'];
        const startColIndex = cols.indexOf(startCol);

        // 1. 줄바꿈을 구분으로 가로 행 분할
        let lines = tsvText.split(/\r?\n/);
        
        // 뒤따르는 불필요한 개행 공백 행이 있으면 필터링
        if (lines.length > 0 && lines[lines.length - 1] === '') {
            lines.pop();
        }

        // 2. 가로 행 내부 데이터들을 탭으로 2차원 배열화
        const dataMatrix = lines.map(line => line.split('\t'));

        // 3. 2차원 루프 순회 돌며 화면 및 비즈니스 데이터 일치 렌더링
        for (let i = 0; i < dataMatrix.length; i++) {
            const targetRow = startRow + i;
            if (targetRow > 9) break; // 9행 경계 보호

            for (let j = 0; j < dataMatrix[i].length; j++) {
                const targetColIndex = startColIndex + j;
                if (targetColIndex > 8) break; // I열 경계 보호

                const targetColLetter = cols[targetColIndex];
                const targetCoord = `${targetColLetter}${targetRow}`;
                const cellValue = dataMatrix[i][j].trim();

                // 화면 td 탐색 및 갱신
                const cellTd = document.querySelector(`.spreadsheet-cell[data-cell="${targetCoord}"]`);
                if (cellTd) {
                    cellTd.textContent = cellValue;
                }

                // 데이터 모델 동시 갱신
                if (cellValue !== '') {
                    this.updateCellValue(targetCoord, cellValue);
                } else {
                    this.deleteCellValue(targetCoord);
                }
            }
        }
    }

    /**
     * [피드백 반영] 시트의 모든 값과 내부 데이터 상태를 완전히 초기화(청소)합니다.
     */
    clearAllSheetData() {
        const cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'];
        for (let r = 1; r <= 9; r++) {
            cols.forEach(c => {
                const cellCoord = `${c}${r}`;
                this.deleteCellValue(cellCoord);
                const td = document.querySelector(`.spreadsheet-cell[data-cell="${cellCoord}"]`);
                if (td) td.textContent = '';
            });
        }
    }

    /**
     * [피드백 반영] One-Click 테스트용 목업 데이터를 A1 원점 기점으로 주입합니다.
     */
    injectSampleMockupData() {
        // [피드백 반영] 예/아니오 대화 경고창 출력 및 유실 방지 가드
        const proceed = confirm("목업 데이터가 시트에 입력되고 현재 데이터는 지워집니다. 입력할까요?");
        if (!proceed) return;

        // [3단계] 싹 비워지기 전, 기존 데이터를 안전하게 되돌리기 위해 미리 백업 저장
        this.saveStateToUndoStack();

        // 기존 모든 시트 데이터를 완벽히 제거 (완전한 오버라이트 주입 사양 충족)
        this.clearAllSheetData();

        const mockupTSV = `품명\t규격\t수량\t단가\t합계\t상태\t구분\t담당자\t비고
MacBook Air\tM3 13"\t12\t1590000\t19080000\t입고완료\tIT자산\t김철수\t영업부 지급용
LG Gram 16\tIntel i7\t8\t1850000\t14800000\t검수중\tIT자산\t이영희\t개발부 지급용
Dell Monitor\tU2723QE\t15\t650000\t9750000\t발주완료\t디스플레이\t박민수\t디자인팀 추가분
MX Master 3S\tLogitech\t20\t139000\t2780000\t입고완료\t소모품\t최재원\t공용 사무용품
Keychron Q1\tGateron\t10\t220000\t2200000\t검수완료\t소모품\t정다은\t프로그래머 전용
iPad Pro 11\tM2 256G\t5\t1240000\t6200000\t출고완료\t태블릿\t강태호\t기획팀 테스트용
Galaxy Tab S9\tUltra\t4\t1370000\t5480000\t입고대기\t태블릿\t윤서연\t모바일개발팀
Office 365\tBusiness\t50\t12100\t605000\t라이선스\t소프트웨어\t한지민\t클라우드 구독`;

        // 4번째 인자(skipUndo)로 true를 전달하여 중복 백업을 건너뜁니다.
        this.importTSVData(mockupTSV, 'A', 1, true);
        
        // 데이터 주입 완료 후 A1 셀을 디폴트 선택 상태로 활성화
        const firstCell = document.querySelector('.spreadsheet-cell[data-cell="A1"]');
        if (firstCell) {
            this.selectCell(firstCell);
        }
        
        console.log("목업 테스트 데이터 주입 성공.");
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
