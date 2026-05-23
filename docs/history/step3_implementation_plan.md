# step3_implementation_plan.md (3단계 구현 계획서 누적 보존본)

본 문서는 사용자의 과제 제출 및 히스토리 증적을 위해 영구 보존되는 3단계 Undo(되돌리기) 시스템 구축의 설계서입니다.

---

## 👨‍💻 추가 기능 및 피드백 설계

### 📌 1. HTML Undo 버튼 UI 배치
* `index.html` 내의 `#sample-btn` 왼쪽에 세련된 디자인의 `Undo ↩️` 버튼을 삽입합니다.

```html
<button id="undo-btn" class="btn btn-secondary">
    Undo ↩️
</button>
```

### 📌 2. `script.js` 내 Undo 데이터 스택 모델링
* **스택 초기화**: `SpreadsheetApp` 클래스의 `constructor()` 내에 `this.undoStack = [];` 멤버 변수를 생성합니다.
* **상태 복제 유틸 (`saveStateToUndoStack()`)**:
  * 데이터 상태가 파괴적이거나 전면 교체되기 **직전**의 `this.spreadsheetData` 객체를 깊은 복사하여 스택에 추가합니다.
  * 메모리 과부하 방지를 위해 스택의 최대 깊이를 **50개**로 고정합니다.
* **상태 변경 전 저장 트리거 배치**:
  1. **값 직접 입력 저장 시 (`exitEditMode`)**: 값이 실제로 변경되었을 때만 트리거.
  2. **일괄 삭제 시 (`clearSelectionValues`)**: 삭제 영역 내 유효 데이터가 존재할 때만 트리거.
  3. **Sample 목업 데이터 주입 시 (`injectSampleMockupData`)**: 청소 전 상태를 1회 백업 저장.
  4. **외부 붙여넣기 주입 시 (`importTSVData`)**: 2차원 데이터를 덮어쓰기 직전에 트리거 (중복 방지 skipUndo 지원).

### 📌 3. `undo()` 실행 취소 구현 및 단축키 (`Ctrl+Z`) 바인딩
* `Undo` 버튼 클릭 혹은 `Ctrl + Z` 입력 시 가장 최근 상태를 `pop()`하여 데이터를 완전 롤백하고, 화면 전체를 청소 후 재렌더링(`renderAllData()`)하며 로컬스토리지를 강제 동기화시킵니다.
* `handleDocumentKeyDown(e)`에서 편집 모드가 아닐 때 `Ctrl + Z`를 가로채 실행 취소를 수행합니다.
