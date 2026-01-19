from scheduler_core import GeneticOptimizer, STAFF_NAMES
from excel_exporter import save_to_excel

def parse_vacation_ranges(vacation_data, num_days):
    """ '이름': [(시작,끝)] -> {(r, c): '휴'} 변환 """
    requests = {}
    for name, ranges in vacation_data.items():
        if name not in STAFF_NAMES: continue
        staff_idx = STAFF_NAMES.index(name)
        for start, end in ranges:
            for day in range(start, end + 1):
                day_idx = day - 1
                if 0 <= day_idx < num_days:
                    requests[(staff_idx, day_idx)] = '휴'
    return requests

def main():
    print("=== 교대근무표 생성 프로그램 시작 ===")
    
    num_staff = 16
    num_days = 31 
    
    # 1. 휴가 일정
    raw_vacation_data = {
        "1팀체계": [(26, 30)],
        "1팀보안": [(26, 31)],
        "2팀체계": [(19, 31)],
        "2팀보안": [],
        "3팀체계": [],
        "3팀보안": [(16, 17)],
        "4팀체계": [(1, 4)],
        "4팀보안": [(1, 1)],
        "체지원1": [(5, 16), (19, 30)],
        "체지원2": [(1, 2)],
        "체지원3": [(1, 8)],
        "체지원4": [(13, 21)],
        "보지원1": [(21, 31)],
        "보지원2": [(13, 19)],
        "보지원3": [(4, 15)],
        "보지원4": [(1, 11), (16, 31)]
    }
    
    vacation_requests = parse_vacation_ranges(raw_vacation_data, num_days)

    # 2. 이번 달 시작 주기 설정 (1일차 근무 형태)
    start_settings = {
        # 정규 팀: 1팀(야) -> 2팀(주) -> 4팀(비) -> 3팀(휴)
        "1팀체계": '야', "1팀보안": '야',
        "2팀체계": '주', "2팀보안": '주',
        "3팀체계": '휴', "3팀보안": '휴',
        "4팀체계": '비', "4팀보안": '비',
        
        # 지원 팀: 정규팀 보조에 맞춰 설정
        "체지원1": '야', "보지원1": '야', 
        "체지원2": '주', "보지원2": '주', 
        "체지원3": '휴', "보지원3": '휴', 
        "체지원4": '비', "보지원4": '비', 
    }
    
    cycle_starts_indices = {}
    pattern_map = {'주':0, '야':1, '비':2, '휴':3} 
    
    for name, start_shift in start_settings.items():
        if name in STAFF_NAMES:
            idx = STAFF_NAMES.index(name)
            cycle_starts_indices[idx] = pattern_map[start_shift]

    # --- 최적화 실행 ---
    print(f"최적의 근무표를 계산 중입니다... (3000세대, 시간이 다소 걸릴 수 있습니다)")
    
    # [변경됨] 세대수 1500 -> 3000
    optimizer = GeneticOptimizer(
        num_staff, 
        num_days, 
        vacation_requests, 
        cycle_starts_indices, 
        pop_size=200, 
        generations=3000 
    )
    
    optimizer.initialize_population()
    
    try:
        best_schedule = optimizer.evolve()
        
        print(f"\n최종 점수: {best_schedule.score:.1f} / 3000")
        
        filename = "2025년_01월_근무표.xlsx"
        save_to_excel(best_schedule, cycle_starts_indices, filename)
        
    except ValueError as e:
        print(f"\n[오류 발생] {e}")

if __name__ == "__main__":
    main()