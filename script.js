/**
 * 객체지향(OOP) 기반 미니 스프레드시트 애플리케이션 (script.js)
 * 
 * [설계 특징]
 * 1. 상태 중심 캡슐화: 전역 변수 대신 SpreadsheetApp 클래스 내부에 데이터와 UI 상태를 완전히 캡슐화.
 * 2. 단일 책임 원칙(SRP): 1함수 1역할을 철저히 고수하여 UI 처리, 데이터 관리, 이벤트 할당, 셀 이동 모듈을 완전 분리.
 * 3. 입문자 친화성: 바이브코딩 초급자도 각 기능의 흐름을 쉽게 이해하도록 모든 핵심 코드 라인에 한글 설명 제공.
 */

class SpreadsheetApp {
    /**
     * 애플리케이션 생성자: 초기 상태 정의 및 엘리먼트 캐싱을 담당합니다.
     */
    constructor() {
        // [데이터 상태 관리] 사용자가 입력한 셀 좌표와 값을 매핑하는 단일 플랫 객체
        this.spreadsheetData = {};
        
        // [DOM 엘리먼트 캐싱] 자주 접근하는 화면 요소들을 변수에 등록해 메모리를 최적화합니다.
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
     * 화면 상의 버튼, 입력창, 헤더들에 각 이벤트를 연결하는 유일한 단일 책임 메서드입니다.
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

        // C. 엑셀 내보내기 버튼 클릭 이벤트
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
     * 셀 포커스 핸들러: 실시간 표시기 갱신 및 헤더 하이라이팅 유도
     */
    handleCellFocus(e) {
        // 기존의 행/열 전체 범위 선택이 있었다면 시각 효과 즉시 해제
        this.clearAllCellSelections();

        const col = e.target.dataset.col;
        const row = e.target.dataset.row;
        const cellCoord = e.target.dataset.cell;

        // UI 업데이트
        this.setIndicatorText(`Cell: ${cellCoord}`, true);
        this.highlightCellHeaders(col, row);
    }

    /**
     * 셀 포커스 이탈 핸들러: 시각 효과 초기화
     */
    handleCellBlur() {
        // 하이라이트 해제 및 기본 텍스트 원복
        this.clearAllHeaderHighlights();
        this.setIndicatorText("Cell: 선택 안 됨", false);
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
                e.preventDefault(); // 엔터 키 기본 동작 차단
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
        if (row < 1 || row > 9) return; // 1행 미만, 9행 초과 무시
        if (col < 'A' || col > 'I') return; // A열 이전, I열 이후 무시

        // 이동할 대상 셀의 Input 태그 탐색
        const targetInput = document.querySelector(`input[data-cell="${col}${row}"]`);
        
        if (targetInput) {
            targetInput.focus();
            // 포커스 진입 시 텍스트 전체 선택 효과를 부여해 엑셀다운 사용성을 높입니다.
            targetInput.select(); 
        }
    }

    /* ==========================================
       [행/열 헤더 클릭 시 전체 행/열 선택 모듈]
       ========================================== */

    /**
     * 열(세로) 헤더 클릭 핸들러: 전체 열 범위 선택
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
    }

    /**
     * 행(가로) 헤더 클릭 핸들러: 전체 행 범위 선택
     */
    handleRowHeaderClick(e) {
        // 1. 기존 하이라이트 초기화
        this.clearAllHeaderHighlights();
        this.clearAllCellSelections();

        // 2. 클릭된 가로 행 번호 파악 (예: "5")
        const row = e.target.dataset.headerRow;

        // 3. 해당 행 헤더 활성화 하이라이트 주입
        e.target.classList.add('active-header');

        // 4. 동일한 행 번호 속성을 지닌 가로라인 <input> 셀들 일괄 선택 효과
        const targetCells = document.querySelectorAll(`input[data-row="${row}"]`);
        targetCells.forEach(cell => {
            cell.classList.add('selected-cell');
        });

        // 5. 좌표 표시기를 해당 행 전체 선택 상태로 갱신
        this.setIndicatorText(`Cell: ${row}행 전체 선택됨`, true);
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
