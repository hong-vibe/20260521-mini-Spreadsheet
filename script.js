/**
 * 미니 스프레드시트 비즈니스 로직 및 이벤트 바인딩 (script.js)
 * 순수 Vanilla JS 기반 설계
 */

// 1. [핵심 데이터 구조] 전체 시트 상태를 단일 플랫 자바스크립트 객체로 관리
// Key: 셀의 좌표 문자열 (예: "A1", "C5"), Value: 사용자가 입력한 값 문자열
const spreadsheetData = {};

// 2. DOM 요소들 탐색 및 캐싱
const currentCellIndicator = document.getElementById('current-cell');
const exportButton = document.getElementById('export-btn');
const cellInputs = document.querySelectorAll('.cell-input');

// 3. 모든 헤더의 하이라이트(.active-header)를 일괄 제거하는 헬퍼 함수
function clearActiveHeaders() {
    const activeHeaders = document.querySelectorAll('.active-header');
    activeHeaders.forEach(header => {
        header.classList.remove('active-header');
    });
}

// 4. 그리드 셀들의 이벤트 바인딩 (focus, blur, input)
cellInputs.forEach(input => {
    
    // [Focus 이벤트]: 셀이 선택되었을 때의 처리
    input.addEventListener('focus', (e) => {
        // 모든 이전 헤더의 하이라이트를 깨끗하게 정리
        clearActiveHeaders();

        // 현재 셀에서 열 알파벳(A~I)과 행 번호(1~9) 데이터 속성 추출
        const col = e.target.dataset.col;
        const row = e.target.dataset.row;
        const cellCoord = e.target.dataset.cell;

        // `#current-cell` 요소의 텍스트를 "Cell: [좌표]" 로 실시간 업데이트
        currentCellIndicator.textContent = `Cell: ${cellCoord}`;
        currentCellIndicator.style.color = 'var(--primary-color)';
        currentCellIndicator.style.borderColor = 'var(--primary-color)';

        // 해당하는 행 헤더와 열 헤더 요소를 각각 DOM에서 탐색
        const colHeader = document.querySelector(`th[data-header-col="${col}"]`);
        const rowHeader = document.querySelector(`th[data-header-row="${row}"]`);

        // 탐색된 헤더 요소에 동적으로 `.active-header` 클래스를 추가하여 강조 표시
        if (colHeader) colHeader.classList.add('active-header');
        if (rowHeader) rowHeader.classList.add('active-header');
    });

    // [Blur 이벤트]: 셀이 포커스를 잃었을 때의 처리
    input.addEventListener('blur', () => {
        // 헤더 하이라이트 클래스를 부드럽게 제거
        clearActiveHeaders();

        // 좌표 인디케이터를 초기 기본 텍스트 상태로 원복
        currentCellIndicator.textContent = "Cell: 선택 안 됨";
        currentCellIndicator.style.color = 'var(--text-muted)';
        currentCellIndicator.style.borderColor = '#e2e8f0';
    });

    // [Input 이벤트]: 셀에 사용자가 데이터를 실시간으로 타이핑할 때의 처리
    input.addEventListener('input', (e) => {
        const cellCoord = e.target.dataset.cell; // A1, C3 등의 셀 좌표 키
        const value = e.target.value.trim();      // 사용자가 입력한 문자열 (좌우 공백 제거)

        if (value !== '') {
            // 현재 입력된 값을 캡처하여 spreadsheetData 객체에 저장
            spreadsheetData[cellCoord] = value;
        } else {
            // [주의] 만약 셀 입력값이 지워져서 빈칸이 되면, spreadsheetData 객체에서 해당 Key를 완전히 삭제(delete)
            delete spreadsheetData[cellCoord];
        }

        // 디버깅용: 콘솔에 현재 스프레드시트 상태를 출력하여 실시간 동기화 상태 검증
        console.log('현재 전역 상태 (spreadsheetData):', spreadsheetData);
    });
});

// 5. [#export-btn] 버튼 클릭 시 SheetJS를 이용해 구글 스프레드시트와 호환되는 엑셀 다운로드 실행
exportButton.addEventListener('click', () => {
    try {
        // SheetJS 워크북 생성
        const wb = XLSX.utils.book_new();

        // 엑셀 시트로 변환할 빈 객체 선언
        const ws = {};

        // [구글 스프레드시트 호환성 극대화 핵심]
        // 9행(1~9) x 9열(A~I)을 아우르는 명시적 범위 설정
        // 이를 설정해주면 비어있는 셀이 있더라도 업로드 시 정확히 A1~I9 그리드가 구글 스프레드시트에 고스란히 생성됩니다.
        ws['!ref'] = "A1:I9";

        // spreadsheetData에 축적된 셀 좌표와 문자열 값들을 SheetJS 셀 구조에 맞추어 매핑
        // 예: ws["A2"] = { v: "내용", t: "s" } (t: 's'는 텍스트 형식을 의미함)
        for (const [cellCoord, val] of Object.entries(spreadsheetData)) {
            ws[cellCoord] = {
                v: val,
                t: 's' // 모든 데이터를 텍스트 포맷으로 매핑하여 입력 유실을 방지합니다.
            };
        }

        // 생성한 시트 객체를 "Sheet1"이라는 이름으로 워크북에 결합
        XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

        // 파일 이름이 "spreadsheet.xlsx"인 실제 엑셀 파일의 다운로드 실행 브라우저 트리거
        XLSX.writeFile(wb, "spreadsheet.xlsx");

        console.log("엑셀 익스포트 성공: spreadsheet.xlsx 파일이 정상적으로 다운로드되었습니다.");
    } catch (error) {
        console.error("엑셀 파일 내보내기 도중 오류가 발생했습니다:", error);
        alert("엑셀 파일 내보내기에 실패했습니다. 다시 시도해 주세요.");
    }
});
