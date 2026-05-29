/**
 * 객체지향(OOP) 기반 미니 스프레드시트 애플리케이션 (script.js)
 * 
 * [ES6 Modules 아키텍처 리팩토링 완료]
 * - SpreadsheetUtils.js : 엑셀 변환, TSV 클립보드 파서, 열-숫자 인덱스 변환기 독립
 * - GameEngine.js : 5단계 스텔스 위장 게임의 타이머, Fisher-Yates 카드 셔플러, 점수 판정 격리
 * - script.js (본 파일) : 화면 그리기(DOM 렌더링)와 마우스/키보드 감지(이벤트 연결) 메인 컨트롤러 역할 전담
 */

import { SpreadsheetUtils } from './SpreadsheetUtils.js';
import { StealthGameEngine } from './GameEngine.js';

class SpreadsheetApp {
    /**
     * 애플리케이션 생성자: 초기 상태 정의 및 엘리먼트 캐싱을 담당합니다.
     */
    constructor() {
        // [데이터 상태 관리] 사용자가 입력한 셀 좌표와 값을 매핑하는 단일 플랫 객체
        this.spreadsheetData = {};

        // [실행 취소(Undo) 상태 관리] 이전의 데이터 모델 상태를 깊은 복사하여 순차적으로 쌓는 히스토리 스택
        this.undoStack = [];

        // [동적 그리드 크기 상태 관리] 4단계 행/열 동적 추가 삭제에 따른 실시간 격자 크기 정보
        this.maxRows = 9;
        this.maxCols = 9;

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

        // [5단계 - 테라피 휴식 게임 전용 상태 변수들 및 부품 엔진]
        this.isTherapyMode = false;         // 현재 테라피 휴식 모드 활성화 상태
        this.gameEngine = null;             // StealthGameEngine 인스턴스 (게임 진입 시 동적 생성)
        this.firstFlippedCard = null;       // 매칭 비교용 첫 번째 클릭된 카드 DOM
        this.secondFlippedCard = null;      // 매칭 비교용 두 번째 클릭된 카드 DOM
        this.lockTherapyBoard = false;      // 0.4초 딜레이 또는 비교 연산 중 클릭 오작동 방지 락
        this.sheetBackup = null;            // 게임 진입 전 시트 데이터 및 크기 대피 백업 컨테이너
        
        // [DOM 엘리먼트 캐싱] 자주 접근하는 화면 요소들을 캐시하여 성능을 극대화합니다.
        this.table = document.querySelector('.spreadsheet-table');
        this.currentCellIndicator = document.getElementById('current-cell');
        this.exportButton = document.getElementById('export-btn');
        this.cells = document.querySelectorAll('.spreadsheet-cell');
        
        // [초기 구동] 이벤트 리스너 등록을 기점으로 애플리케이션을 구동시킵니다.
        this.init();
    }

    /**
     * 초기화 모듈: 상태 복구 및 동적 그리드 리빌드를 거쳐 이벤트를 장착합니다.
     */
    init() {
        this.loadFromLocalStorage(); // [2단계] 로컬스토리지에서 기존 데이터 및 격자 크기 정보 선 복구 로드
        this.rebuildGrid();          // [4단계] 획득한 크기 사양에 맞춰 DOM 그리드 완전 리빌드 및 리사이저 결합
        this.bindEvents();           // 전역 컨트롤 버튼 및 단축키 등 이벤트 바인딩
    }

    /**
     * [이벤트 처리 모듈]
     * 화면 상의 버튼, 셀, 헤더들에 각 이벤트를 연결하는 단일 책임 메서드입니다.
     */
    bindEvents() {
        // A. 개별 셀(td)들에 대한 이벤트 연결
        this.cells.forEach(cell => {
            cell.addEventListener('mousedown', (e) => this.handleCellMouseDown(e));
            cell.addEventListener('mouseenter', (e) => this.handleCellMouseEnter(e));
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

        // E. 좌측 상단 모서리(.corner-header) 클릭 시 전체 선택 처리 바인딩
        const cornerHeader = document.querySelector('.corner-header');
        if (cornerHeader) {
            cornerHeader.addEventListener('click', () => this.handleCornerHeaderClick());
        }

        // F. 원클릭 목업 주입 Sample 버튼 이벤트 바인딩
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

        // I. 앱 전면 초기화(Reset) 버튼 클릭 이벤트 바인딩
        const resetBtn = document.getElementById('reset-btn');
        if (resetBtn) {
            resetBtn.addEventListener('click', () => this.resetSheetToDefault());
        }

        // J. 엑셀 내보내기 버튼 클릭 이벤트
        this.exportButton.addEventListener('click', () => this.exportToExcel());

        // K. [5단계] 테라피 휴식 게임 버튼 클릭 이벤트
        const therapyBtn = document.getElementById('therapy-btn');
        if (therapyBtn) {
            therapyBtn.addEventListener('click', () => this.toggleTherapyMode());
        }
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
     * 현재 spreadsheetData의 데이터 상태를 깊은 복사하여 Undo 스택에 저장합니다.
     * 메모리 누수를 방지하기 위해 최대 스택 크기를 50개로 제한합니다.
     */
    saveStateToUndoStack() {
        const snapshot = {
            data: JSON.parse(JSON.stringify(this.spreadsheetData)),
            maxRows: this.maxRows,
            maxCols: this.maxCols
        };
        this.undoStack.push(snapshot);

        if (this.undoStack.length > 50) {
            this.undoStack.shift();
        }
        console.log('실행 취소(Undo) 스택 저장 완료. 현재 크기:', this.undoStack.length);
    }

    /**
     * 가장 최근에 저장했던 상태로 롤백하고 화면과 로컬스토리지를 강제 동기화합니다.
     */
    undo() {
        if (this.isEditing && this.currentSelection && this.currentSelection.element) {
            this.exitEditMode(this.currentSelection.element, false);
        }

        if (this.undoStack.length === 0) {
            console.log('되돌릴 이전 이력이 존재하지 않습니다.');
            return;
        }

        const prevState = this.undoStack.pop();
        this.spreadsheetData = prevState.data;
        this.maxRows = prevState.maxRows;
        this.maxCols = prevState.maxCols;

        // 과거 크기에 맞춰 화면 그리드 DOM 전면 리빌드 및 데이터 자동 주입
        this.rebuildGrid();

        this.saveToLocalStorage();

        this.currentSelection = null;
        this.setIndicatorText("Cell: 선택 안 됨");
        this.clearAllHeaderHighlights();

        console.log('실행 취소(Undo) 완료. 남은 스택 크기:', this.undoStack.length);
    }

    /* ==========================================
       [2단계 - 로컬스토리지 자동 저장 및 복구 엔진 모듈]
       ========================================== */

    /**
     * 현재의 전역 스프레드시트 데이터 상태 및 격자 크기를 로컬스토리지에 오토 세이브합니다.
     */
    saveToLocalStorage() {
        try {
            localStorage.setItem('pingpong_spreadsheet_data', JSON.stringify(this.spreadsheetData));
            localStorage.setItem('pingpong_spreadsheet_max_rows', this.maxRows);
            localStorage.setItem('pingpong_spreadsheet_max_cols', this.maxCols);
            console.log(`로컬스토리지 자동 저장 성공. 크기: ${this.maxRows} x ${this.maxCols}`);
        } catch (error) {
            console.error('로컬스토리지 저장 실패:', error);
        }
    }

    /**
     * 페이지 첫 기동 시 로컬스토리지를 검사해 기존 데이터 상태와 격자 크기를 완벽 복구 로드합니다.
     */
    loadFromLocalStorage() {
        try {
            const savedRows = localStorage.getItem('pingpong_spreadsheet_max_rows');
            const savedCols = localStorage.getItem('pingpong_spreadsheet_max_cols');
            if (savedRows) this.maxRows = parseInt(savedRows, 10);
            if (savedCols) this.maxCols = parseInt(savedCols, 10);

            const saved = localStorage.getItem('pingpong_spreadsheet_data');
            if (saved) {
                this.spreadsheetData = JSON.parse(saved);
                console.log(`로컬스토리지 복구 감지. 크기: ${this.maxRows} x ${this.maxCols}`, this.spreadsheetData);
            } else {
                console.log('로컬스토리지에 저장된 이전 세션 데이터가 존재하지 않습니다.');
            }
        } catch (error) {
            console.error('로컬스토리지 복구 처리 중 장애 발생:', error);
        }
    }

    /**
     * 복구된 spreadsheetData 상태를 루프하여 화면 그리드 td 엘리먼트에 채워줍니다.
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
        if (this.isTherapyMode) return;
        if (this.isEditing && this.currentSelection && this.currentSelection.element === e.currentTarget) {
            return;
        }

        if (this.isEditing && this.currentSelection && this.currentSelection.element) {
            this.exitEditMode(this.currentSelection.element, true);
        }

        const td = e.currentTarget;
        const col = td.dataset.col;
        const row = parseInt(td.dataset.row, 10);

        if (e.shiftKey) {
            e.preventDefault();
            this.handleShiftClickSelection(td);
            return;
        }

        this.selectCell(td);
        this.isDragging = true;
        this.dragStartCell = { col, row };
        this.dragEndCell = { col, row };
    }

    /**
     * 셀 마우스 엔터 핸들러: 드래그 활성화 중일 때, 마우스 궤적에 따른 영역을 하이라이트합니다.
     */
    handleCellMouseEnter(e) {
        if (this.isTherapyMode) return;
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

        if (start.col === end.col && start.row === end.row) {
            return;
        }

        this.currentSelection = {
            type: 'range',
            startCol: start.col,
            startRow: start.row,
            endCol: end.col,
            endRow: end.row,
            element: this.currentSelection.element
        };

        // 실시간 인디케이터에 선택된 사각형 영역 크기를 표현
        const cols = [];
        for (let c = 0; c < this.maxCols; c++) {
            cols.push(SpreadsheetUtils.getColLetter(c));
        }
        const sColIdx = cols.indexOf(start.col);
        const eColIdx = cols.indexOf(end.col);
        const colCount = Math.abs(sColIdx - eColIdx) + 1;
        const rowCount = Math.abs(start.row - end.row) + 1;
        
        this.setIndicatorText(`Cell: ${colCount}R x ${rowCount}C`, true);
    }

    /**
     * Shift + 클릭 선택 핸들러: 기존 포커싱 원점을 기준으로 사각형 영역을 계산해 일괄 선택합니다.
     */
    handleShiftClickSelection(targetTd) {
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

        this.updateDragSelection();

        this.currentSelection = {
            type: 'range',
            startCol: startCol,
            startRow: startRow,
            endCol: endCol,
            endRow: endRow,
            element: anchorTd
        };

        const cols = [];
        for (let c = 0; c < this.maxCols; c++) {
            cols.push(SpreadsheetUtils.getColLetter(c));
        }
        const sColIdx = cols.indexOf(startCol);
        const eColIdx = cols.indexOf(endCol);
        const colCount = Math.abs(sColIdx - eColIdx) + 1;
        const rowCount = Math.abs(startRow - endRow) + 1;

        this.setIndicatorText(`Cell: ${colCount}R x ${rowCount}C`, true);
    }

    /**
     * dragStartCell과 dragEndCell 사이의 사각형 영역 내 모든 td에 하이라이트를 실시간 업데이트합니다.
     */
    updateDragSelection() {
        const start = this.dragStartCell;
        const end = this.dragEndCell;

        if (!start || !end) return;

        // 현재 활성화된 컬럼 목록 생성
        const cols = [];
        for (let c = 0; c < this.maxCols; c++) {
            cols.push(SpreadsheetUtils.getColLetter(c));
        }

        const sColIdx = cols.indexOf(start.col);
        const eColIdx = cols.indexOf(end.col);

        const minColIdx = Math.min(sColIdx, eColIdx);
        const maxColIdx = Math.max(sColIdx, eColIdx);
        const minRow = Math.min(start.row, end.row);
        const maxRow = Math.max(start.row, end.row);

        this.clearAllCellSelections();

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
        if (this.isTherapyMode) return;
        const td = e.currentTarget;
        this.enterEditMode(td);
    }

    /**
     * 특정 셀을 '선택 대기' 상태(엑셀 초록 테두리)로 전환하는 내부 보조 메서드입니다.
     */
    selectCell(tdElement) {
        this.clearAllCellSelections();
        this.clearSelectedCellClass();

        const col = tdElement.dataset.col;
        const row = parseInt(tdElement.dataset.row, 10);
        const cellCoord = tdElement.dataset.cell;

        tdElement.classList.add('cell-selected');

        this.setIndicatorText(`Cell: ${cellCoord}`, true);
        this.highlightCellHeaders(col, row);

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
     */
    enterEditMode(tdElement, initialValue = null) {
        if (this.isEditing) return;
        this.isEditing = true;

        tdElement.classList.add('cell-editing-active');

        const cellCoord = tdElement.dataset.cell;
        const originalVal = this.spreadsheetData[cellCoord] || '';

        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'cell-input-edit';

        if (initialValue !== null) {
            input.value = initialValue;
        } else {
            input.value = originalVal;
        }

        tdElement.textContent = '';
        tdElement.appendChild(input);
        
        input.focus();

        if (initialValue === null) {
            input.select();
        }

        input.addEventListener('keydown', (e) => {
            const col = tdElement.dataset.col;
            const row = parseInt(tdElement.dataset.row, 10);

            if (e.key === 'Enter') {
                e.preventDefault();
                this.exitEditMode(tdElement, true);
                this.moveFocus(col, row + 1);
            } else if (e.key === 'Escape') {
                e.preventDefault();
                this.exitEditMode(tdElement, false);
                this.selectCell(tdElement);
            } else if (e.key === 'Tab') {
                e.preventDefault();
                this.exitEditMode(tdElement, true);
                const nextCol = String.fromCharCode(col.charCodeAt(0) + 1);
                this.moveFocus(nextCol, row);
            }
        });

        input.addEventListener('blur', () => {
            this.exitEditMode(tdElement, true);
        });
    }

    /**
     * [수정 모드 이탈]: 동적 input을 제거하고 원래의 td 텍스트 렌더링으로 롤백/저장합니다.
     */
    exitEditMode(tdElement, shouldSave) {
        if (!this.isEditing) return;

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
            if (finalValue !== originalVal) {
                this.saveStateToUndoStack();
            }

            if (finalValue !== '') {
                this.updateCellValue(cellCoord, finalValue);
                tdElement.textContent = finalValue;
            } else {
                this.deleteCellValue(cellCoord);
                tdElement.textContent = '';
            }
        } else {
            tdElement.textContent = this.spreadsheetData[cellCoord] || '';
        }

        this.isEditing = false;
    }

    /* ==========================================
       [키보드 네비게이션 및 전역 덮어쓰기 감지 모듈]
       ========================================== */

    /**
     * 전역 도큐먼트 키다운 이벤트 분기 처리
     */
    handleDocumentKeyDown(e) {
        if (this.isTherapyMode) return;
        
        if (e.ctrlKey && (e.key === 'z' || e.key === 'Z')) {
            e.preventDefault();
            this.undo();
            return;
        }

        if (this.isEditing) return;
        if (!this.currentSelection) return;

        if (e.key === 'Delete' || e.key === 'Backspace') {
            e.preventDefault();
            this.clearSelectionValues();
            return;
        }

        if (this.currentSelection.type !== 'cell') return;

        const col = this.currentSelection.col;
        const row = this.currentSelection.row;
        const currentTd = this.currentSelection.element;

        switch (e.key) {
            case 'Enter':
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

        if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
            e.preventDefault();
            this.enterEditMode(currentTd, e.key);
        }
    }

    /**
     * 특정 좌표로 포커스를 즉시 안전 이동시키는 제어 함수입니다.
     */
    moveFocus(col, row) {
        if (row < 1 || row > this.maxRows) return; 
        const colIdx = SpreadsheetUtils.getColIndex(col);
        if (colIdx < 0 || colIdx >= this.maxCols) return; 

        const targetTd = document.querySelector(`.spreadsheet-cell[data-cell="${col}${row}"]`);
        if (targetTd) {
            this.selectCell(targetTd);
        }
    }

    /**
     * 현재 선택 범위의 셀 텍스트 및 전역 데이터를 일괄 제거합니다.
     */
    clearSelectionValues() {
        if (!this.currentSelection) return;

        const cols = [];
        for (let c = 0; c < this.maxCols; c++) {
            cols.push(SpreadsheetUtils.getColLetter(c));
        }

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
                for (let r = 1; r <= this.maxRows; r++) {
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

        if (hasDataToDelete) {
            this.saveStateToUndoStack();
        }

        switch (this.currentSelection.type) {
            case 'cell':
                const coord = this.currentSelection.cell;
                this.deleteCellValue(coord);
                if (this.currentSelection.element) {
                    this.currentSelection.element.textContent = '';
                }
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
                        const colLetter = cols[cIdx];
                        const cellCoord = `${colLetter}${r}`;
                        this.deleteCellValue(cellCoord);
                        const td = document.querySelector(`.spreadsheet-cell[data-cell="${cellCoord}"]`);
                        if (td) td.textContent = '';
                    }
                }
                break;

            case 'col':
                const cLetter = this.currentSelection.col;
                for (let r = 1; r <= this.maxRows; r++) {
                    const cellCoord = `${cLetter}${r}`;
                    this.deleteCellValue(cellCoord);
                    const td = document.querySelector(`.spreadsheet-cell[data-cell="${cellCoord}"]`);
                    if (td) td.textContent = '';
                }
                break;

            case 'row':
                const rNum = this.currentSelection.row;
                cols.forEach(c => {
                    const cellCoord = `${c}${rNum}`;
                    this.deleteCellValue(cellCoord);
                    const td = document.querySelector(`.spreadsheet-cell[data-cell="${cellCoord}"]`);
                    if (td) td.textContent = '';
                });
                break;

            case 'all':
                for (let r = 1; r <= this.maxRows; r++) {
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
       [열 가로 너비(Width) 조절 마우스 조작 모듈]
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

            handle.addEventListener('mousedown', (e) => this.handleResizeMouseDown(e, th));
        });
    }

    /**
     * 리사이즈 드래그 시작 mousedown
     */
    handleResizeMouseDown(e, th) {
        e.preventDefault();
        e.stopPropagation();

        this.isResizing = true;
        this.resizingTh = th;
        this.resizeStartX = e.clientX;
        this.resizeStartWidth = th.offsetWidth;

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
        const newWidth = Math.max(50, this.resizeStartWidth + deltaX);

        this.resizingTh.style.width = `${newWidth}px`;
    }

    /**
     * 리사이즈 드래그 종료 mouseup
     */
    handleResizeMouseUp() {
        if (!this.isResizing) return;
        this.isResizing = false;
        this.resizingTh = null;

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
        if (e.target.classList.contains('resize-handle')) return;

        this.clearAllHeaderHighlights();
        this.clearAllCellSelections();
        this.clearSelectedCellClass();

        const col = e.target.dataset.headerCol;
        e.target.classList.add('active-header');

        const targetCells = document.querySelectorAll(`.spreadsheet-cell[data-col="${col}"]`);
        targetCells.forEach(cell => {
            cell.classList.add('selected-cell');
        });

        this.setIndicatorText(`Cell: ${col}열`, true);

        this.currentSelection = {
            type: 'col',
            col: col
        };
    }

    /**
     * 행(가로) 헤더 클릭 핸들러: 전체 행 범위 선택 및 복사 범위 지정
     */
    handleRowHeaderClick(e) {
        this.clearAllHeaderHighlights();
        this.clearAllCellSelections();
        this.clearSelectedCellClass();

        const row = parseInt(e.target.dataset.headerRow, 10);
        e.target.classList.add('active-header');

        const targetCells = document.querySelectorAll(`.spreadsheet-cell[data-row="${row}"]`);
        targetCells.forEach(cell => {
            cell.classList.add('selected-cell');
        });

        this.setIndicatorText(`Cell: ${row}행`, true);

        this.currentSelection = {
            type: 'row',
            row: row
        };
    }

    /**
     * 모서리(코너) 최상단 좌측 셀 클릭 핸들러: 전체 영역 선택 및 복사 지정
     */
    handleCornerHeaderClick() {
        this.clearAllHeaderHighlights();
        this.clearAllCellSelections();
        this.clearSelectedCellClass();

        const allThHeaders = document.querySelectorAll('.spreadsheet-table th:not(.corner-header)');
        allThHeaders.forEach(th => {
            th.classList.add('active-header');
        });

        this.cells.forEach(cell => {
            cell.classList.add('selected-cell');
        });

        this.setIndicatorText("Cell: ALL", true);

        this.currentSelection = {
            type: 'all'
        };
    }

    /* ==========================================
       [양방향 클립보드 복사 / 붙여넣기 전용 모듈]
       ========================================== */

    /**
     * 복사(Ctrl+C) 이벤트 가로채기 핸들러
     */
    handleClipboardCopy(e) {
        if (this.isTherapyMode) return;
        if (!this.currentSelection) return;

        let copyText = '';
        
        const cols = [];
        for (let c = 0; c < this.maxCols; c++) {
            cols.push(SpreadsheetUtils.getColLetter(c));
        }

        switch (this.currentSelection.type) {
            case 'cell':
                const coord = this.currentSelection.cell;
                copyText = this.spreadsheetData[coord] || '';
                break;

            case 'range':
                const sColIdx = cols.indexOf(this.currentSelection.startCol);
                const eColIdx = cols.indexOf(this.currentSelection.endCol);
                const minColIdx = Math.min(sColIdx, eColIdx);
                const maxColIdx = Math.max(sColIdx, eColIdx);
                
                const minRow = Math.min(this.currentSelection.startRow, this.currentSelection.endRow);
                const maxRow = Math.max(this.currentSelection.startRow, this.currentSelection.endRow);
                
                const rangeRows = [];
                for (let r = minRow; r <= maxRow; r++) {
                    const rowCells = [];
                    for (let cIdx = minColIdx; cIdx <= maxColIdx; cIdx++) {
                        const colLetter = cols[cIdx];
                        const cellCoord = `${colLetter}${r}`;
                        rowCells.push(this.spreadsheetData[cellCoord] || '');
                    }
                    rangeRows.push(rowCells.join('\t'));
                }
                copyText = rangeRows.join('\r\n');
                break;

            case 'col':
                const colLetter = this.currentSelection.col;
                const colValues = [];
                for (let r = 1; r <= this.maxRows; r++) {
                    const cCoord = `${colLetter}${r}`;
                    colValues.push(this.spreadsheetData[cCoord] || '');
                }
                copyText = colValues.join('\r\n');
                break;

            case 'row':
                const rowNum = this.currentSelection.row;
                const rowValues = [];
                cols.forEach(c => {
                    const rCoord = `${c}${rowNum}`;
                    rowValues.push(this.spreadsheetData[rCoord] || '');
                });
                copyText = rowValues.join('\t');
                break;

            case 'all':
                const gridRows = [];
                for (let r = 1; r <= this.maxRows; r++) {
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

        e.clipboardData.setData('text/plain', copyText);
        e.preventDefault();
        
        console.log("클립보드에 스프레드시트 호환 복사 완료:", copyText);
    }

    /**
     * 붙여넣기(Ctrl+V) 이벤트 가로채기 핸들러
     */
    handleClipboardPaste(e) {
        if (this.isTherapyMode) return;
        if (!this.currentSelection || this.currentSelection.type !== 'cell') return;

        const clipboardData = e.clipboardData || window.clipboardData;
        const pastedText = clipboardData.getData('text');
        
        if (!pastedText) return;

        const startCol = this.currentSelection.col;
        const startRow = this.currentSelection.row;

        e.preventDefault();

        this.importTSVData(pastedText, startCol, startRow);
        console.log("외부 클립보드 다중 셀 2차원 붙여넣기 착지 완료.");
    }

    /**
     * 탭(TSV)으로 구분된 데이터를 좌표에 맞춰 2차원 해독 렌더링하고 상태 모델에 기입하는 핵심 공통 모듈입니다.
     */
    importTSVData(tsvText, startCol, startRow, skipUndo = false) {
        if (!skipUndo) {
            this.saveStateToUndoStack();
        }

        // [부품 위임] 탭 문자 해독 처리를 SpreadsheetUtils에 전격 위임합니다.
        const dataMatrix = SpreadsheetUtils.parseTSV(tsvText);

        const cols = [];
        for (let c = 0; c < this.maxCols; c++) {
            cols.push(SpreadsheetUtils.getColLetter(c));
        }
        const startColIndex = cols.indexOf(startCol);

        for (let i = 0; i < dataMatrix.length; i++) {
            const targetRow = startRow + i;
            if (targetRow > this.maxRows) break; // 동적 행 경계 보호

            for (let j = 0; j < dataMatrix[i].length; j++) {
                const targetColIndex = startColIndex + j;
                if (targetColIndex >= this.maxCols) break; // 동적 열 경계 보호

                const targetColLetter = cols[targetColIndex];
                const targetCoord = `${targetColLetter}${targetRow}`;
                const cellValue = dataMatrix[i][j].trim();

                const cellTd = document.querySelector(`.spreadsheet-cell[data-cell="${targetCoord}"]`);
                if (cellTd) {
                    cellTd.textContent = cellValue;
                }

                if (cellValue !== '') {
                    this.updateCellValue(targetCoord, cellValue);
                } else {
                    this.deleteCellValue(targetCoord);
                }
            }
        }
    }

    /**
     * 시트의 모든 값과 내부 데이터 상태를 완전히 초기화(청소)합니다.
     */
    clearAllSheetData() {
        const cols = [];
        for (let c = 0; c < this.maxCols; c++) {
            cols.push(SpreadsheetUtils.getColLetter(c));
        }

        for (let r = 1; r <= this.maxRows; r++) {
            cols.forEach(c => {
                const cellCoord = `${c}${r}`;
                this.deleteCellValue(cellCoord);
                const td = document.querySelector(`.spreadsheet-cell[data-cell="${cellCoord}"]`);
                if (td) td.textContent = '';
            });
        }
    }

    /**
     * One-Click 테스트용 목업 데이터를 A1 원점 기점으로 주입합니다.
     */
    injectSampleMockupData() {
        const proceed = confirm("목업 데이터가 시트에 입력되고 현재 데이터는 지워집니다. 입력할까요?");
        if (!proceed) return;

        this.saveStateToUndoStack();
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

        this.importTSVData(mockupTSV, 'A', 1, true);
        
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
     * [부품 위임] Excel 변환 처리를 SpreadsheetUtils에 전격 이관하여 실행합니다.
     */
    exportToExcel() {
        try {
            SpreadsheetUtils.exportToExcel(this.spreadsheetData, this.maxCols, this.maxRows);
            console.log("엑셀 다운로드 처리 완료: spreadsheet.xlsx가 로컬로 저장되었습니다.");
        } catch (error) {
            console.error("SheetJS 파일 내보내기 장애 발생:", error);
            alert("엑셀 파일 변환 도중 예상치 못한 오류가 발생했습니다.");
        }
    }

    /**
     * 시트를 완전히 초기 상태인 9x9 빈 격자로 원복하고 로컬스토리지를 깨끗이 비웁니다.
     */
    resetSheetToDefault() {
        const proceed = confirm("시트가 초기화 되고 현재 데이터는 지워집니다. 초기화 할까요?");
        if (!proceed) return;

        this.saveStateToUndoStack();

        this.spreadsheetData = {};
        this.maxRows = 9;
        this.maxCols = 9;

        localStorage.removeItem('pingpong_spreadsheet_data');
        localStorage.removeItem('pingpong_spreadsheet_max_rows');
        localStorage.removeItem('pingpong_spreadsheet_max_cols');

        this.rebuildGrid();
        this.saveToLocalStorage();

        this.currentSelection = null;
        this.setIndicatorText("Cell: 선택 안 됨");
        this.clearAllHeaderHighlights();

        console.log("미니 스프레드시트가 9x9 원점 상태로 완전 초기화 완료되었습니다.");
    }

    /* ==========================================
       [4단계 - 행/열 동적 추가 및 삭제 모듈 - 헤더 플로팅 단추 연동식]
       ========================================== */

    /**
     * 현재 maxRows와 maxCols 정보에 맞춰 DOM 테이블 격자를 실시간 재생성하고 데이터를 주입합니다.
     */
    rebuildGrid() {
        if (!this.table) return;

        const thead = this.table.querySelector('thead');
        const tbody = this.table.querySelector('tbody');
        if (!thead || !tbody) return;

        thead.innerHTML = '';
        tbody.innerHTML = '';

        const headerTr = document.createElement('tr');
        const cornerTh = document.createElement('th');
        cornerTh.className = 'corner-header';
        headerTr.appendChild(cornerTh);

        for (let c = 0; c < this.maxCols; c++) {
            const colLetter = SpreadsheetUtils.getColLetter(c);
            const th = document.createElement('th');
            th.setAttribute('data-header-col', colLetter);
            
            const textNode = document.createTextNode(colLetter);
            th.appendChild(textNode);

            const btnGroup = document.createElement('div');
            btnGroup.className = 'header-btn-group';

            const addBtn = document.createElement('button');
            addBtn.className = 'mini-btn mini-btn-add';
            addBtn.textContent = '+';
            addBtn.title = `${colLetter}열 우측에 새 열 추가`;
            addBtn.addEventListener('click', (evt) => {
                evt.stopPropagation();
                this.addCol(colLetter);
            });

            const delBtn = document.createElement('button');
            delBtn.className = 'mini-btn mini-btn-delete';
            delBtn.textContent = '-';
            delBtn.title = `${colLetter}열 삭제`;
            delBtn.addEventListener('click', (evt) => {
                evt.stopPropagation();
                this.deleteCol(colLetter);
            });

            btnGroup.appendChild(addBtn);
            btnGroup.appendChild(delBtn);
            th.appendChild(btnGroup);

            headerTr.appendChild(th);
        }
        thead.appendChild(headerTr);

        for (let r = 1; r <= this.maxRows; r++) {
            const rowTr = document.createElement('tr');
            
            const rowTh = document.createElement('th');
            rowTh.setAttribute('data-header-row', r);
            
            const textNode = document.createTextNode(r);
            rowTh.appendChild(textNode);

            const btnGroup = document.createElement('div');
            btnGroup.className = 'header-btn-group';

            const addBtn = document.createElement('button');
            addBtn.className = 'mini-btn mini-btn-add';
            addBtn.textContent = '+';
            addBtn.title = `${r}행 아래쪽에 새 행 추가`;
            addBtn.addEventListener('click', (evt) => {
                evt.stopPropagation();
                this.addRow(r);
            });

            const delBtn = document.createElement('button');
            delBtn.className = 'mini-btn mini-btn-delete';
            delBtn.textContent = '-';
            delBtn.title = `${r}행 삭제`;
            delBtn.addEventListener('click', (evt) => {
                evt.stopPropagation();
                this.deleteRow(r);
            });

            btnGroup.appendChild(addBtn);
            btnGroup.appendChild(delBtn);
            rowTh.appendChild(btnGroup);

            rowTr.appendChild(rowTh);

            for (let c = 0; c < this.maxCols; c++) {
                const colLetter = SpreadsheetUtils.getColLetter(c);
                const cellCoord = `${colLetter}${r}`;
                const td = document.createElement('td');
                td.className = 'spreadsheet-cell';
                td.setAttribute('data-cell', cellCoord);
                td.setAttribute('data-col', colLetter);
                td.setAttribute('data-row', r);

                td.textContent = this.spreadsheetData[cellCoord] || '';
                rowTr.appendChild(td);
            }
            tbody.appendChild(rowTr);
        }

        this.cells = document.querySelectorAll('.spreadsheet-cell');
        this.initColumnResizers();
        this.rebindCellEvents();
        this.rebindHeaderEvents();
    }

    /**
     * 동적으로 교체된 셀(td)들에 이벤트 리스너를 다시 바인딩합니다.
     */
    rebindCellEvents() {
        this.cells.forEach(cell => {
            cell.addEventListener('mousedown', (e) => this.handleCellMouseDown(e));
            cell.addEventListener('mouseenter', (e) => this.handleCellMouseEnter(e));
            cell.addEventListener('dblclick', (e) => this.handleCellDblClick(e));
        });
    }

    /**
     * 동적으로 교체된 헤더(th)들에 이벤트 리스너를 재결합합니다.
     */
    rebindHeaderEvents() {
        const columnHeaders = this.table.querySelectorAll('th[data-header-col]');
        const rowHeaders = this.table.querySelectorAll('th[data-header-row]');

        columnHeaders.forEach(th => {
            th.addEventListener('click', (e) => {
                if (e.target.classList.contains('mini-btn') || e.target.closest('.header-btn-group')) return;
                this.handleColumnHeaderClick(e);
            });
        });

        rowHeaders.forEach(th => {
            th.addEventListener('click', (e) => {
                if (e.target.classList.contains('mini-btn') || e.target.closest('.header-btn-group')) return;
                this.handleRowHeaderClick(e);
            });
        });

        const cornerHeader = this.table.querySelector('.corner-header');
        if (cornerHeader) {
            cornerHeader.addEventListener('click', () => this.handleCornerHeaderClick());
        }
    }

    /**
     * 특정 행 아래에 새로운 행을 추가하고 데이터 좌표를 아래로 1칸씩 밀어냅니다 (Shift Down).
     */
    addRow(targetRow = null) {
        if (targetRow === null) {
            if (this.currentSelection) {
                if (this.currentSelection.type === 'cell') {
                    targetRow = this.currentSelection.row;
                } else if (this.currentSelection.type === 'range') {
                    targetRow = this.currentSelection.startRow;
                } else if (this.currentSelection.type === 'row') {
                    targetRow = this.currentSelection.row;
                }
            } else {
                targetRow = this.maxRows;
            }
        }

        targetRow = parseInt(targetRow, 10);
        const insertAt = targetRow + 1;

        this.saveStateToUndoStack();

        const newData = {};
        for (const [key, value] of Object.entries(this.spreadsheetData)) {
            const col = key.match(/[A-Z]+/)[0];
            const r = parseInt(key.match(/[0-9]+/)[0], 10);
            if (r >= insertAt) {
                newData[`${col}${r + 1}`] = value;
            } else {
                newData[key] = value;
            }
        }
        this.spreadsheetData = newData;

        this.maxRows++;
        this.rebuildGrid();
        this.saveToLocalStorage();

        console.log(`행 추가 완료. 삽입 위치: ${insertAt}행, 전체 행수: ${this.maxRows}`);
    }

    /**
     * 지정된 행을 삭제하고 아래쪽의 데이터 좌표를 위로 1칸씩 당깁니다 (Shift Up).
     */
    deleteRow(targetRow = null) {
        if (this.maxRows <= 1) {
            alert("최소 1개의 행이 존재해야 합니다.");
            return;
        }

        if (targetRow === null) {
            if (this.currentSelection) {
                if (this.currentSelection.type === 'cell') {
                    targetRow = this.currentSelection.row;
                } else if (this.currentSelection.type === 'range') {
                    targetRow = this.currentSelection.startRow;
                } else if (this.currentSelection.type === 'row') {
                    targetRow = this.currentSelection.row;
                }
            } else {
                targetRow = this.maxRows;
            }
        }

        targetRow = parseInt(targetRow, 10);

        this.saveStateToUndoStack();

        const newData = {};
        for (const [key, value] of Object.entries(this.spreadsheetData)) {
            const col = key.match(/[A-Z]+/)[0];
            const r = parseInt(key.match(/[0-9]+/)[0], 10);
            if (r === targetRow) {
                continue;
            } else if (r > targetRow) {
                newData[`${col}${r - 1}`] = value;
            } else {
                newData[key] = value;
            }
        }
        this.spreadsheetData = newData;

        this.maxRows = Math.max(1, this.maxRows - 1);
        
        this.currentSelection = null;
        this.setIndicatorText("Cell: 선택 안 됨");
        this.clearAllHeaderHighlights();

        this.rebuildGrid();
        this.saveToLocalStorage();

        console.log(`행 삭제 완료. 대상: ${targetRow}행, 전체 행수: ${this.maxRows}`);
    }

    /**
     * 지정한 열 우측에 새로운 열을 추가하고 데이터 좌표를 우측으로 1칸씩 밀어냅니다 (Shift Right).
     */
    addCol(targetColLetter = null) {
        let targetColIdx = this.maxCols - 1;
        if (targetColLetter !== null) {
            targetColIdx = SpreadsheetUtils.getColIndex(targetColLetter);
        } else if (this.currentSelection) {
            if (this.currentSelection.type === 'cell') {
                targetColIdx = SpreadsheetUtils.getColIndex(this.currentSelection.col);
            } else if (this.currentSelection.type === 'range') {
                targetColIdx = SpreadsheetUtils.getColIndex(this.currentSelection.startCol);
            } else if (this.currentSelection.type === 'col') {
                targetColIdx = SpreadsheetUtils.getColIndex(this.currentSelection.col);
            }
        }

        const insertAtIdx = targetColIdx + 1;

        this.saveStateToUndoStack();

        const newData = {};
        for (const [key, value] of Object.entries(this.spreadsheetData)) {
            const col = key.match(/[A-Z]+/)[0];
            const r = parseInt(key.match(/[0-9]+/)[0], 10);
            const cIdx = SpreadsheetUtils.getColIndex(col);
            
            if (cIdx >= insertAtIdx) {
                const nextColLetter = SpreadsheetUtils.getColLetter(cIdx + 1);
                newData[`${nextColLetter}${r}`] = value;
            } else {
                newData[key] = value;
            }
        }
        this.spreadsheetData = newData;

        this.maxCols++;
        this.rebuildGrid();
        this.saveToLocalStorage();

        console.log(`열 추가 완료. 삽입 위치: ${SpreadsheetUtils.getColLetter(insertAtIdx)}열, 전체 열수: ${this.maxCols}`);
    }

    /**
     * 지정한 열을 삭제하고 우측의 데이터 좌표를 좌측으로 1칸씩 당깁니다 (Shift Left).
     */
    deleteCol(targetColLetter = null) {
        if (this.maxCols <= 1) {
            alert("최소 1개의 열이 존재해야 합니다.");
            return;
        }

        let targetColIdx = this.maxCols - 1;
        if (targetColLetter !== null) {
            targetColIdx = SpreadsheetUtils.getColIndex(targetColLetter);
        } else if (this.currentSelection) {
            if (this.currentSelection.type === 'cell') {
                targetColIdx = SpreadsheetUtils.getColIndex(this.currentSelection.col);
            } else if (this.currentSelection.type === 'range') {
                targetColIdx = SpreadsheetUtils.getColIndex(this.currentSelection.startCol);
            } else if (this.currentSelection.type === 'col') {
                targetColIdx = SpreadsheetUtils.getColIndex(this.currentSelection.col);
            }
        }

        this.saveStateToUndoStack();

        const newData = {};
        for (const [key, value] of Object.entries(this.spreadsheetData)) {
            const col = key.match(/[A-Z]+/)[0];
            const r = parseInt(key.match(/[0-9]+/)[0], 10);
            const cIdx = SpreadsheetUtils.getColIndex(col);

            if (cIdx === targetColIdx) {
                continue;
            } else if (cIdx > targetColIdx) {
                const prevColLetter = SpreadsheetUtils.getColLetter(cIdx - 1);
                newData[`${prevColLetter}${r}`] = value;
            } else {
                newData[key] = value;
            }
        }
        this.spreadsheetData = newData;

        this.maxCols = Math.max(1, this.maxCols - 1);

        this.currentSelection = null;
        this.setIndicatorText("Cell: 선택 안 됨");
        this.clearAllHeaderHighlights();

        this.rebuildGrid();
        this.saveToLocalStorage();

        console.log(`열 삭제 완료. 대상 인덱스: ${targetColIdx}, 전체 열수: ${this.maxCols}`);
    }

    /* ==========================================
       [5단계 - 오피스 스텔스 테라피 휴식 매칭 게임 독립 라이프사이클 엔진 결합]
       ========================================== */

    /**
     * 테라피 휴식 게임 모드의 온/오프 상태를 제어하는 토글러입니다.
     */
    toggleTherapyMode() {
        if (this.isEditing && this.currentSelection && this.currentSelection.element) {
            this.exitEditMode(this.currentSelection.element, true);
        }

        if (this.isTherapyMode) {
            this.stopTherapyMode();
        } else {
            this.startTherapyMode();
        }
    }

    /**
     * 테라피 휴식 모드 개시
     */
    startTherapyMode() {
        this.isTherapyMode = true;

        this.clearAllCellSelections();
        this.clearSelectedCellClass();
        this.clearAllHeaderHighlights();

        this.sheetBackup = {
            data: JSON.parse(JSON.stringify(this.spreadsheetData)),
            rows: this.maxRows,
            cols: this.maxCols
        };

        // 기존 액션 버튼들 visually disable 처리
        this.exportButton.style.opacity = '0.5';
        this.exportButton.style.pointerEvents = 'none';
        const sampleBtn = document.getElementById('sample-btn');
        if (sampleBtn) {
            sampleBtn.style.opacity = '0.5';
            sampleBtn.style.pointerEvents = 'none';
        }
        const undoBtn = document.getElementById('undo-btn');
        if (undoBtn) {
            undoBtn.style.opacity = '0.5';
            undoBtn.style.pointerEvents = 'none';
        }
        const resetBtn = document.getElementById('reset-btn');
        if (resetBtn) {
            resetBtn.style.opacity = '0.5';
            resetBtn.style.pointerEvents = 'none';
        }

        const therapyBtn = document.getElementById('therapy-btn');
        if (therapyBtn) {
            therapyBtn.classList.add('toolbar-btn-active');
            therapyBtn.setAttribute('title', '자유시간 메모리 게임 종료');
        }

        // 은밀한 대시보드 동적 삽입
        const appContainer = document.querySelector('.app-container');
        const gridContainer = document.querySelector('.grid-container');
        const dashboard = document.createElement('div');
        dashboard.className = 'therapy-dashboard';
        dashboard.id = 'therapy-dashboard';

        // [부품 위임] 임시 엔진 인스턴스를 만들고 최고 기록을 가져옵니다.
        // 엔진 생성 시, 1초 타이머가 흘렀을 때 화면 글자(00:00)를 바꾸는 전원 연결부 콜백을 주입합니다.
        this.gameEngine = new StealthGameEngine((elapsedSeconds) => {
            this.updateTherapyTimerDisplay(elapsedSeconds);
        });

        const bestRecordStr = this.gameEngine.getBestRecord() || '없음';
        dashboard.innerHTML = `
            <div>☕ STEALTH ANALYZER - A1:D5 DATA RANGE SELECTED</div>
            <div>
                최고 기록: <span class="best-value">${bestRecordStr}</span> | 
                경과 시간: <span class="timer-value" id="therapy-timer-val">00:00</span>
            </div>
        `;

        appContainer.insertBefore(dashboard, gridContainer);

        this.initTherapyGame();
    }

    /**
     * 테라피 휴식 모드 철수: 원래 셀 데이터로 완전 복원합니다.
     */
    stopTherapyMode() {
        this.isTherapyMode = false;

        // [부품 위임] 타이머 소멸 처리 위임
        if (this.gameEngine) {
            this.gameEngine.stopTimer();
            this.gameEngine = null;
        }

        const dashboard = document.getElementById('therapy-dashboard');
        if (dashboard) {
            dashboard.remove();
        }

        const successPanel = document.getElementById('stealth-success-panel');
        if (successPanel) {
            successPanel.remove();
        }

        this.exportButton.style.opacity = '1';
        this.exportButton.style.pointerEvents = 'auto';
        const sampleBtn = document.getElementById('sample-btn');
        if (sampleBtn) {
            sampleBtn.style.opacity = '1';
            sampleBtn.style.pointerEvents = 'auto';
        }
        const undoBtn = document.getElementById('undo-btn');
        if (undoBtn) {
            undoBtn.style.opacity = '1';
            undoBtn.style.pointerEvents = 'auto';
        }
        const resetBtn = document.getElementById('reset-btn');
        if (resetBtn) {
            resetBtn.style.opacity = '1';
            resetBtn.style.pointerEvents = 'auto';
        }

        const therapyBtn = document.getElementById('therapy-btn');
        if (therapyBtn) {
            therapyBtn.classList.remove('toolbar-btn-active');
            therapyBtn.setAttribute('title', '자유시간 메모리 게임 시작');
        }

        // A1~D5의 20개 셀의 백업 텍스트 복구
        const cols = ['A', 'B', 'C', 'D'];
        for (let r = 1; r <= 5; r++) {
            cols.forEach(c => {
                const coord = `${c}${r}`;
                const td = document.querySelector(`.spreadsheet-cell[data-cell="${coord}"]`);
                if (td) {
                    const originalText = td.getAttribute('data-original-val') || '';
                    td.textContent = originalText;
                    
                    td.removeAttribute('data-original-val');
                    td.removeAttribute('data-flag');
                    td.className = 'spreadsheet-cell';
                    td.style.padding = '';
                }
            });
        }

        if (this.sheetBackup) {
            this.spreadsheetData = this.sheetBackup.data;
            this.maxRows = this.sheetBackup.rows;
            this.maxCols = this.sheetBackup.cols;
        }
        this.rebuildGrid();

        this.setIndicatorText("Cell: 선택 안 됨");
    }

    /**
     * A1~D5 영역에 위장 카드 테두리를 씌우고 실물 국기 매핑 렌더링
     */
    initTherapyGame() {
        this.firstFlippedCard = null;
        this.secondFlippedCard = null;
        this.lockTherapyBoard = false;

        // [부품 위임] 엔진에게 Fisher-Yates 셔플링된 20개 국기 카드를 섞어오라고 시킵니다.
        const shuffledFlags = this.gameEngine.generateShuffledCards();

        const targetCoords = [];
        const cols = ['A', 'B', 'C', 'D'];
        for (let r = 1; r <= 5; r++) {
            cols.forEach(c => {
                targetCoords.push(`${c}${r}`);
            });
        }

        // 20개 td 셀들을 스텔스 카드로 리빌드
        targetCoords.forEach((coord, idx) => {
            const td = document.querySelector(`.spreadsheet-cell[data-cell="${coord}"]`);
            if (td) {
                const flagCode = shuffledFlags[idx];
                const originalVal = td.textContent;

                td.setAttribute('data-original-val', originalVal);
                td.setAttribute('data-flag', flagCode);
                td.className = 'spreadsheet-cell stealth-game-cell';

                td.innerHTML = `
                    <div class="stealth-card-inner">
                        <div class="stealth-card-front">${originalVal}</div>
                        <div class="stealth-card-back">
                            <img src="https://flagcdn.com/w40/${flagCode.toLowerCase()}.png" width="32" height="24" style="object-fit:cover; border-radius:2px;">
                        </div>
                    </div>
                `;

                td.addEventListener('click', (evt) => {
                    evt.stopPropagation();
                    this.handleCardClick(td);
                });
            }
        });
    }

    /**
     * 카드 클릭 핸들러: 플립 모션 작동 및 400ms 기민한 짝 맞추기 평가
     */
    handleCardClick(cardEl) {
        if (cardEl.classList.contains('matched')) return;

        if (this.lockTherapyBoard || cardEl.classList.contains('flipped')) {
            cardEl.classList.add('stealth-shake');
            setTimeout(() => {
                cardEl.classList.remove('stealth-shake');
            }, 300);
            return;
        }

        cardEl.classList.add('flipped');

        // [부품 위임] 첫 조작 시 엔진 시계 개시
        this.gameEngine.startTimer();

        if (!this.firstFlippedCard) {
            this.firstFlippedCard = cardEl;
        } else {
            this.secondFlippedCard = cardEl;
            this.lockTherapyBoard = true;

            const firstFlag = this.firstFlippedCard.getAttribute('data-flag');
            const secondFlag = this.secondFlippedCard.getAttribute('data-flag');

            // [부품 위임] 매칭 판정을 엔진에 전격 위임합니다.
            const isMatch = this.gameEngine.checkMatch(firstFlag, secondFlag);

            if (isMatch) {
                this.firstFlippedCard.classList.add('matched');
                this.secondFlippedCard.classList.add('matched');

                this.firstFlippedCard = null;
                this.secondFlippedCard = null;
                this.lockTherapyBoard = false;

                // 10쌍 올 클리어 도달 시 성공 모듈 구동
                if (this.gameEngine.matchedCount === 10) {
                    this.handleTherapyGameClear();
                }
            } else {
                // 불일치 시 400ms 전격 회전 단축 복귀
                setTimeout(() => {
                    this.firstFlippedCard.classList.remove('flipped');
                    this.secondFlippedCard.classList.remove('flipped');

                    this.firstFlippedCard = null;
                    this.secondFlippedCard = null;
                    this.lockTherapyBoard = false;
                }, 400);
            }
        }
    }

    /**
     * 대시보드 경과 시간 00:00 분:초 단위 디스플레이 포맷팅
     */
    updateTherapyTimerDisplay(elapsedSeconds) {
        const timerVal = document.getElementById('therapy-timer-val');
        if (!timerVal) return;

        const mins = Math.floor(elapsedSeconds / 60);
        const secs = elapsedSeconds % 60;
        timerVal.textContent = `${mins.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}`;
    }

    /**
     * 게임 올 클리어 연출 및 로컬스토리지 완료 레코드 영구 마킹
     */
    handleTherapyGameClear() {
        // [부품 위임] 타이머 멈춤
        this.gameEngine.stopTimer();

        let elapsedText = '';
        if (this.gameEngine.timeElapsed < 60) {
            elapsedText = `${this.gameEngine.timeElapsed}초`;
        } else {
            const m = Math.floor(this.gameEngine.timeElapsed / 60);
            const s = this.gameEngine.timeElapsed % 60;
            elapsedText = `${m}분 ${s}초`;
        }

        // [부품 위임] 점수 오토세이브 위임
        const result = this.gameEngine.saveRecord(elapsedText);

        const appContainer = document.querySelector('.app-container');
        const successDiv = document.createElement('div');
        successDiv.className = 'stealth-success-panel';
        successDiv.id = 'stealth-success-panel';
        successDiv.innerHTML = `
            <div class="success-title">🎉 리프레싱 스텔스 완수! 🎉</div>
            <p style="color: #64748b; font-size: 0.85rem; margin-bottom:0.75rem;">A1:D5 업무 위장 범위의 모든 국기 짝을 정확히 맞추고 두뇌 회전을 끝마쳤습니다.</p>
            
            <div class="success-stats">
                <strong>완료 일시:</strong> ${result ? result.dateStr : ''} ${result ? result.timeStr : ''}<br>
                <strong>소요 시간:</strong> <span style="color: #107c41; font-weight:700;">${elapsedText}</span>
            </div>

            <button id="therapy-close-btn" class="btn btn-primary" style="padding: 0.5rem 2rem; font-size:0.95rem; height:38px; width:100%;">시트로 돌아가기 🏓</button>
            
            <div class="record-board" id="record-board-container"></div>
        `;

        appContainer.appendChild(successDiv);

        // 명예의 전당 보드판 렌더링
        this.renderRecordBoard();

        const closeBtn = document.getElementById('therapy-close-btn');
        if (closeBtn) {
            closeBtn.addEventListener('click', () => this.toggleTherapyMode());
        }
    }

    /**
     * 명예의 전당 기록 리스트를 화면에 렌더링합니다. (순수 DOM 작업)
     */
    renderRecordBoard() {
        const container = document.getElementById('record-board-container');
        if (!container) return;

        try {
            // [부품 위임] 저장된 전체 랭킹 목록을 엔진에서 꺼내옵니다.
            const records = this.gameEngine.getRecords();

            if (records.length === 0) {
                container.innerHTML = `<div style="text-align:center; color:#a855f7; font-size:0.75rem;">아직 수립된 기록이 없습니다.</div>`;
                return;
            }

            let html = `
                <div class="record-board-title">
                    🏆 명예의 전당 (Top 5 Best Times)
                </div>
            `;

            records.slice(0, 5).forEach((rec, idx) => {
                const rankIcon = idx === 0 ? '🥇' : idx === 1 ? '🥈' : idx === 2 ? '🥉' : `${idx + 1}위`;
                html += `
                    <div class="record-item">
                        <span>
                            <span class="record-rank">${rankIcon}</span> 
                            (${rec.date} ${rec.time})
                        </span>
                        <strong style="color: #7e22ce;">${rec.elapsed}</strong>
                    </div>
                `;
            });

            container.innerHTML = html;
        } catch (e) {
            container.innerHTML = `<div style="color:red; font-size:0.75rem;">기록판 파싱 중 장애가 발생했습니다.</div>`;
        }
    }
}

// [애플리케이션 즉시 실행] DOM 로드가 끝나면 바로 클래스 객체를 생성하여 인스턴스화합니다.
document.addEventListener('DOMContentLoaded', () => {
    window.spreadsheetApp = new SpreadsheetApp();
});
