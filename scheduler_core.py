import random
import copy
import numpy as np

# --- 상수 및 설정 정의 ---
SHIFTS = ['출', '주', '당', '야', '생', '휴']
DAY_SHIFTS = ['출', '주']
NIGHT_SHIFTS = ['당', '야']
WORKING_SHIFTS = DAY_SHIFTS + NIGHT_SHIFTS
OFF_SHIFTS = ['생', '휴']

# 이상적인 4일 주기 (주 -> 야 -> 비 -> 생)
CYCLE_PATTERN = ['주', '야', '생', '생']

STAFF_NAMES = [
    "1팀체계", "1팀보안", 
    "2팀체계", "2팀보안", 
    "3팀체계", "3팀보안", 
    "4팀체계", "4팀보안", 
    "체지원1", "체지원2", "체지원3", "체지원4",
    "보지원1", "보지원2", "보지원3", "보지원4"
]

SYSTEM_STAFF_IDX = [i for i, name in enumerate(STAFF_NAMES) if '체' in name]
SECURITY_STAFF_IDX = [i for i, name in enumerate(STAFF_NAMES) if '보' in name]
TEAM_MEMBER_IDX = [i for i, name in enumerate(STAFF_NAMES) if '팀' in name and '지원' not in name]

class Schedule:
    def __init__(self, num_staff, num_days, requests=None, cycle_starts=None, grid=None):
        self.num_staff = num_staff
        self.num_days = num_days
        self.requests = requests if requests else {}
        self.cycle_starts = cycle_starts if cycle_starts else {}
        self.score = 0
        
        if grid is None:
            self.grid = self._create_valid_grid()
        else:
            self.grid = grid

    def _create_valid_grid(self):
        """
        [스마트 초기화 강화판]
        1. 파트너 사이클 우선 배정
        2. [중요] 내일 휴가인 사람은 오늘 절대 '야간' 근무 불가 처리
        3. 지원조 리더 우선 배정
        """
        grid = [[None for _ in range(self.num_days)] for _ in range(self.num_staff)]
        
        for day in range(self.num_days):
            candidates_day = []
            candidates_night = []
            candidates_off = []
            
            # 1. 후보군 분류
            for staff_idx in range(self.num_staff):
                # 이미 휴가 신청된 날이면 고정
                if (staff_idx, day) in self.requests:
                    grid[staff_idx][day] = '휴'
                    continue
                
                # [핵심] 내일 휴가인지 확인
                # 내일 휴가라면 오늘 '야간' 근무는 절대 불가 -> 강제로 Day나 Off로 분류
                is_tomorrow_vacation = False
                if day < self.num_days - 1:
                    if (staff_idx, day + 1) in self.requests:
                        is_tomorrow_vacation = True

                if staff_idx in self.cycle_starts:
                    start_offset = self.cycle_starts[staff_idx]
                    cycle_char = CYCLE_PATTERN[(day + start_offset) % 4]
                    
                    if cycle_char == '주':
                        candidates_day.append(staff_idx)
                    elif cycle_char == '야':
                        # 원래 야간 순서지만 내일 휴가라면? -> 야간 후보 박탈
                        if is_tomorrow_vacation:
                            candidates_off.append(staff_idx) # 쉬거나 주간으로
                        else:
                            candidates_night.append(staff_idx)
                    else:
                        candidates_off.append(staff_idx)
                else:
                    candidates_off.append(staff_idx)

            # 2. 주간 조(3명) 선발
            selected_day_staff = []
            random.shuffle(candidates_day)
            selected_day_staff.extend(candidates_day)
            
            # 부족하면 휴무조 충원
            if len(selected_day_staff) < 3:
                random.shuffle(candidates_off)
                while len(selected_day_staff) < 3 and candidates_off:
                    selected_day_staff.append(candidates_off.pop(0))
            
            # 그래도 부족하면 야간조에서 충원 (극단적 상황)
            if len(selected_day_staff) < 3:
                # 여기서도 내일 휴가인 사람은 야간 후보에서 제외했으므로 안전하지만 한 번 더 필터링
                safe_night_candidates = [x for x in candidates_night if not ((x, day+1) in self.requests)]
                random.shuffle(safe_night_candidates)
                while len(selected_day_staff) < 3 and safe_night_candidates:
                    staff = safe_night_candidates.pop(0)
                    selected_day_staff.append(staff)
                    candidates_night.remove(staff) # 야간 후보에서 제거
            
            selected_day_staff = selected_day_staff[:3]

            # 3. 야간 조(3명) 선발
            selected_night_staff = []
            
            # 남은 야간 후보 중 선발 (이미 내일 휴가자는 걸러져 있음)
            remaining_night = [x for x in candidates_night if x not in selected_day_staff]
            random.shuffle(remaining_night)
            selected_night_staff.extend(remaining_night)
            
            # 부족하면 남은 휴무조 충원
            if len(selected_night_staff) < 3:
                remaining_off = [x for x in candidates_off if x not in selected_day_staff]
                
                # [안전장치] 휴무조에서 데려올 때도 내일 휴가자는 제외해야 함
                # (혹시나 candidates_off에 내일 휴가자가 섞여 있을 수 있으므로)
                safe_off_candidates = []
                for s in remaining_off:
                    if day < self.num_days - 1 and (s, day+1) in self.requests:
                        continue # 내일 휴가면 야간 대타 불가
                    safe_off_candidates.append(s)
                
                random.shuffle(safe_off_candidates)
                while len(selected_night_staff) < 3 and safe_off_candidates:
                    selected_night_staff.append(safe_off_candidates.pop(0))
            
            selected_night_staff = selected_night_staff[:3]

            # --- 역할 배정 (지원조 리더 우선) ---
            def sort_key(idx):
                name = STAFF_NAMES[idx]
                if '지원' in name: return 0
                return 1

            # 주간 배정
            selected_day_staff.sort(key=sort_key)
            if len(selected_day_staff) > 0: grid[selected_day_staff[0]][day] = '출'
            if len(selected_day_staff) > 1: grid[selected_day_staff[1]][day] = '주'
            if len(selected_day_staff) > 2: grid[selected_day_staff[2]][day] = '주'

            # 야간 배정
            selected_night_staff.sort(key=sort_key)
            if len(selected_night_staff) > 0: grid[selected_night_staff[0]][day] = '당'
            if len(selected_night_staff) > 1: grid[selected_night_staff[1]][day] = '야'
            if len(selected_night_staff) > 2: grid[selected_night_staff[2]][day] = '야'
            
            # 나머지 인원 '생'
            for idx in range(self.num_staff):
                if grid[idx][day] is None:
                    grid[idx][day] = '생'

        return grid

    def mutate(self, mutation_rate=0.1):
        new_grid = copy.deepcopy(self.grid)
        for day in range(self.num_days):
            if random.random() < mutation_rate:
                swappable_indices = [i for i in range(self.num_staff) if new_grid[i][day] != '휴']
                if len(swappable_indices) >= 2:
                    idx_a, idx_b = random.sample(swappable_indices, 2)
                    new_grid[idx_a][day], new_grid[idx_b][day] = \
                    new_grid[idx_b][day], new_grid[idx_a][day]
        return Schedule(self.num_staff, self.num_days, self.requests, self.cycle_starts, new_grid)

class Evaluator:
    def evaluate(self, schedule):
        score = 5000 # 기본 점수 대폭 상향
        
        # 1. [CRITICAL] 직능 균형
        score -= self._check_role_balance(schedule) * 500
        
        # 2. [FATAL] 야간 근무 후 휴식 (절대 규칙)
        # 야간 다음날 '휴'가 오면 점수를 마이너스로 보내버릴 정도로 강력하게 응징
        score -= self._check_rest_after_night(schedule) * 50000 
        
        # 3. [STRICT] 정규 팀 파트너 사이클 준수
        score -= self._check_cycle_compliance(schedule) * 500 
        
        # 4. 리더 근무 우선권
        score -= self._check_leader_priority(schedule) * 300

        # 5. 연속 근무 제한
        score -= self._check_progressive_consecutive_work(schedule) * 100

        # 6. 연속 휴무 제한
        score -= self._check_excessive_consecutive_off(schedule) * 50

        # 7. 근무 시간 형평성
        score -= self._check_working_hours_fairness(schedule) * 5

        return score

    def _check_leader_priority(self, schedule):
        penalty = 0
        grid = schedule.grid
        for d in range(schedule.num_days):
            # 주간
            day_indices = [i for i in range(schedule.num_staff) if grid[i][d] in DAY_SHIFTS]
            if day_indices:
                leader_idx = next((i for i in day_indices if grid[i][d] == '출'), None)
                if leader_idx is not None:
                    leader_is_support = '지원' in STAFF_NAMES[leader_idx]
                    has_support_member = any('지원' in STAFF_NAMES[i] for i in day_indices)
                    if has_support_member and not leader_is_support:
                        penalty += 1
            # 야간
            night_indices = [i for i in range(schedule.num_staff) if grid[i][d] in NIGHT_SHIFTS]
            if night_indices:
                leader_idx = next((i for i in night_indices if grid[i][d] == '당'), None)
                if leader_idx is not None:
                    leader_is_support = '지원' in STAFF_NAMES[leader_idx]
                    has_support_member = any('지원' in STAFF_NAMES[i] for i in night_indices)
                    if has_support_member and not leader_is_support:
                        penalty += 1
        return penalty

    def _check_role_balance(self, schedule):
        penalty = 0
        grid = schedule.grid
        for d in range(schedule.num_days):
            day_group = []
            night_group = []
            for s in range(schedule.num_staff):
                shift = grid[s][d]
                if shift in DAY_SHIFTS: day_group.append(s)
                elif shift in NIGHT_SHIFTS: night_group.append(s)
            
            if not (any(idx in SYSTEM_STAFF_IDX for idx in day_group) and 
                    any(idx in SECURITY_STAFF_IDX for idx in day_group)):
                penalty += 1
            if not (any(idx in SYSTEM_STAFF_IDX for idx in night_group) and 
                    any(idx in SECURITY_STAFF_IDX for idx in night_group)):
                penalty += 1
        return penalty

    def _check_rest_after_night(self, schedule):
        """
        야간('당', '야') 다음날은 무조건 '생'이어야 함.
        '휴'가 오는 것을 절대적으로 막기 위해 페널티를 매우 크게 부여.
        """
        penalty = 0
        for r in range(schedule.num_staff):
            for c in range(schedule.num_days - 1):
                prev = schedule.grid[r][c]
                curr = schedule.grid[r][c+1]
                
                if prev in NIGHT_SHIFTS:
                    # 다음날 '생'이 아니면 모두 위반
                    # 특히 '휴'인 경우도 포함됨
                    if curr != '생': 
                        penalty += 1
        return penalty

    def _check_cycle_compliance(self, schedule):
        penalty = 0
        for r in range(schedule.num_staff):
            if r not in TEAM_MEMBER_IDX: continue
            if r not in schedule.cycle_starts: continue
            
            start_offset = schedule.cycle_starts[r]
            for c in range(schedule.num_days):
                if schedule.grid[r][c] == '휴': continue
                
                expected_type = CYCLE_PATTERN[(c + start_offset) % 4]
                actual_shift = schedule.grid[r][c]
                
                is_match = False
                if expected_type == '주' and actual_shift in DAY_SHIFTS: is_match = True
                elif expected_type == '야' and actual_shift in NIGHT_SHIFTS: is_match = True
                elif expected_type == '생' and actual_shift == '생': is_match = True
                
                if not is_match: 
                    penalty += 1
        return penalty

    def _check_progressive_consecutive_work(self, schedule):
        penalty = 0
        for row in schedule.grid:
            consecutive = 0
            for shift in row:
                if shift in WORKING_SHIFTS:
                    consecutive += 1
                else:
                    consecutive = 0
                
                if consecutive == 3: penalty += 1 
                elif consecutive == 4: penalty += 5
                elif consecutive >= 5: penalty += 10
        return penalty

    def _check_excessive_consecutive_off(self, schedule):
        penalty = 0
        limit = 2
        for row in schedule.grid:
            consecutive = 0
            for shift in row:
                if shift in OFF_SHIFTS:
                    consecutive += 1
                else:
                    consecutive = 0
                if consecutive > limit: penalty += 1
        return penalty

    def _check_working_hours_fairness(self, schedule):
        total_hours = []
        for r in range(schedule.num_staff):
            hours = 0
            for shift in schedule.grid[r]:
                if shift in DAY_SHIFTS: hours += 8
                elif shift in NIGHT_SHIFTS: hours += 13
            total_hours.append(hours)
        return np.std(total_hours)

class GeneticOptimizer:
    def __init__(self, num_staff, num_days, requests, cycle_starts, pop_size=50, generations=100):
        self.num_staff = num_staff
        self.num_days = num_days
        self.requests = requests
        self.cycle_starts = cycle_starts
        self.pop_size = pop_size
        self.generations = generations
        self.evaluator = Evaluator()
        self.population = []

    def initialize_population(self):
        for _ in range(self.pop_size):
            self.population.append(Schedule(self.num_staff, self.num_days, self.requests, self.cycle_starts))

    def evolve(self):
        best_schedule = None
        for gen in range(self.generations):
            for individual in self.population:
                individual.score = self.evaluator.evaluate(individual)
            self.population.sort(key=lambda x: x.score, reverse=True)
            
            if best_schedule is None or self.population[0].score > best_schedule.score:
                best_schedule = copy.deepcopy(self.population[0])

            if gen % 50 == 0:
                print(f"[알고리즘 진행중] 세대 {gen}: 점수 = {self.population[0].score:.1f}")
            
            if self.population[0].score >= 4900:
                print(">>> 최적해 발견! <<<")
                break

            num_elites = int(self.pop_size * 0.2)
            next_generation = self.population[:num_elites]
            while len(next_generation) < self.pop_size:
                parent = random.choice(self.population[:num_elites])
                child = parent.mutate(mutation_rate=0.2) 
                next_generation.append(child)
            self.population = next_generation
        return best_schedule