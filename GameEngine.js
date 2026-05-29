/**
 * GameEngine.js
 * 
 * [비밀 게임 브레인]
 * 5단계 오피스 스텔스 국기 매치 ☕ 테라피 휴식 게임 엔진
 * 
 * - 카드 셔플 (Fisher-Yates 알고리즘)
 * - 게임 타이머(1초 단위 비동기 똑딱이) 및 경과 시간 상태 관리
 * - 카드 짝 맞춤 판정 심판 기능
 * - 명예의 전당 점수판 오토세이브(LocalStorage 연동 및 초 단위 랭킹 정렬)
 * 
 * 이 파일 역시 순수한 논리 연산과 게임 상태를 관리하는 엔진으로, DOM을 만지는 화면 렌더링은 메인 조율실로 분리합니다.
 */

export class StealthGameEngine {
    /**
     * 게임 엔진 생성자
     * @param {Function} onTimerTick - 1초 단위 시간 경과를 메인 화면(UI)에 전달하는 통신용 콜백 배달원
     */
    constructor(onTimerTick) {
        this.timeElapsed = 0;        // 흘러간 총 시간(초 단위)
        this.timerInterval = null;   // setInterval 타이머 인스턴스 객체
        this.matchedCount = 0;       // 현재 성공적으로 맞춘 짝 개수 (최대 10개)
        this.onTimerTick = onTimerTick; // 1초마다 실행될 갱신 피드백
        this.cards = [];             // 셔플되어 배치된 20개 국기 리스트
    }

    /**
     * 평화 친화적 10개국 국기 카드를 무작위 셔플링하여 20장(2쌍씩)의 카드를 생성합니다.
     * Fisher-Yates 무작위 카드 섞기 마술 알고리즘 적용.
     * 
     * @returns {Array<string>} 셔플된 20개 국기 코드 배열 (KR, FR, DE, GB 등)
     */
    generateShuffledCards() {
        const baseFlags = ['KR', 'FR', 'DE', 'GB', 'CA', 'BR', 'IT', 'ES', 'AU', 'CH'];
        this.cards = [...baseFlags, ...baseFlags];
        
        // 뒤쪽 카드부터 하나씩 앞으로 가며 무작위로 위치 교환
        for (let i = this.cards.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [this.cards[i], this.cards[j]] = [this.cards[j], this.cards[i]];
        }
        
        this.matchedCount = 0;
        this.timeElapsed = 0;
        return this.cards;
    }

    /**
     * 힐링 경과 똑딱이 타이머 시계를 작동합니다.
     */
    startTimer() {
        // 이미 타이머가 가동 중이라면 중복 생성 억제
        if (this.timerInterval) return;

        this.timeElapsed = 0;
        this.timerInterval = setInterval(() => {
            this.timeElapsed++;
            // 메인 UI 화면에게 흘러간 초 단위를 알려줌
            if (this.onTimerTick) {
                this.onTimerTick(this.timeElapsed);
            }
        }, 1000);
    }

    /**
     * 힐링 타이머 시계를 정지하고 파괴합니다.
     */
    stopTimer() {
        if (this.timerInterval) {
            clearInterval(this.timerInterval);
            this.timerInterval = null;
        }
    }

    /**
     * 두 뒤집힌 카드의 국기가 일치하는지 심판을 내립니다.
     * 일치할 경우 획득 짝수를 증가시키고 true를 반환합니다.
     * 
     * @param {string} flag1 - 첫 번째 카드의 국기 코드
     * @param {string} flag2 - 두 번째 카드의 국기 코드
     * @returns {boolean} 일치 여부
     */
    checkMatch(flag1, flag2) {
        if (flag1 === flag2) {
            this.matchedCount++;
            return true;
        }
        return false;
    }

    /**
     * 게임 클리어 시 완료 날짜, 시:분 정보 및 소요시간 랭킹 레코드를 브라우저 전용 로컬 저장창고에 적재합니다.
     * 
     * @param {string} elapsedText - 예쁘게 포맷팅된 시간 문자열 (예: "35초" 또는 "1분 5초")
     * @returns {Object|null} 저장 완료 데이터 패키지 (dateStr, timeStr, records)
     */
    saveRecord(elapsedText) {
        try {
            const now = new Date();
            
            // YYYY-MM-DD 포맷
            const dateStr = `${now.getFullYear()}-${(now.getMonth() + 1).toString().padStart(2, '0')}-${now.getDate().toString().padStart(2, '0')}`;
            
            // HH:MM (시:분만 표시 사양)
            const timeStr = `${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}`;

            // 기존 명예의 전당 명단 가져오기
            const rawRecords = localStorage.getItem('pingpong_therapy_records');
            let records = rawRecords ? JSON.parse(rawRecords) : [];

            // 새 레코드 삽입
            records.push({
                date: dateStr,
                time: timeStr,
                elapsed: elapsedText,
                seconds: this.timeElapsed
            });

            // 소요시간(초)이 가장 적은 순으로 정렬하여 명예의 전당 탑 랭킹 수립
            records.sort((a, b) => a.seconds - b.seconds);

            // 최대 10개 기록까지만 보존 관리
            records = records.slice(0, 10);

            // 로컬스토리지 영구 적재
            localStorage.setItem('pingpong_therapy_records', JSON.stringify(records));

            return {
                dateStr,
                timeStr,
                records
            };
        } catch (e) {
            console.error("로컬스토리지 명예의 전당 기록 갱신 실패:", e);
            return null;
        }
    }

    /**
     * 역대 1위 최고 기록(Best Record) 스트링 획득
     * 
     * @returns {string|null} 최단 시간 소요값 텍스트 또는 null
     */
    getBestRecord() {
        try {
            const rawRecords = localStorage.getItem('pingpong_therapy_records');
            if (rawRecords) {
                const records = JSON.parse(rawRecords);
                if (records.length > 0) {
                    return records[0].elapsed; // 최단 시간 소요값 반환
                }
            }
        } catch (e) {}
        return null;
    }

    /**
     * 현재 로컬스토리지에 저장된 명예의 전당 랭킹 데이터 전체 조회
     * 
     * @returns {Array<Object>} 랭킹 전체 레코드 리스트
     */
    getRecords() {
        try {
            const rawRecords = localStorage.getItem('pingpong_therapy_records');
            return rawRecords ? JSON.parse(rawRecords) : [];
        } catch (e) {
            console.error("랭킹 데이터 파싱 에러:", e);
            return [];
        }
    }
}
