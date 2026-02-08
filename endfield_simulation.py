import random
import time
import xlsxwriter
import statistics

# ==========================================

# 模拟次数 (建议测试时用100，正式跑时根据自己需求决定)
SIMULATION_COUNT = 100000

# 初始抽数为120
initial_permit_default = 120

annual_permit_update_default = 200

# 基础参数
daily_permit_update = 200 * 21 / 500
weekly_permit_update = 500 * 3 / 500
season_pass_permit_update = (600 + 3 * 75) / 1000
custom_pass_permit_update = (75 * 36) / 1000
monthly_pass_permit_update = (200 * 21 + 75 * 12 / 30 * 14) / 500
calendar_permit_update = 7.5
custom_pass_ars_update = 2400

free_permit_update = daily_permit_update + weekly_permit_update + season_pass_permit_update + calendar_permit_update
monthly_permit_update = free_permit_update + monthly_pass_permit_update
custom_permit_update = monthly_permit_update + custom_pass_permit_update

free_ars_update = 300
custom_ars_update = free_ars_update + custom_pass_ars_update / 2

# 输出Sheet配置
CONFIGS = {
    "TIER_0 (0氪)": {
        "init": initial_permit_default,
        "update": free_permit_update,
        "ars": free_ars_update,
        "annual": 0,
        "sheet_name": "0氪"
    },
    "TIER_1 (月卡)": {
        "init": initial_permit_default,
        "update": monthly_permit_update,
        "ars": free_ars_update,
        "annual": 0,
        "sheet_name": "月卡"
    },
    "TIER_2 (大月卡)": {
        "init": initial_permit_default,
        "update": custom_permit_update,
        "ars": custom_ars_update,
        "annual": 0,
        "sheet_name": "大小月卡"
    },
    "TIER_3 (首充双倍)": {
        "init": initial_permit_default,
        "update": custom_permit_update,
        "ars": custom_ars_update,
        "annual": annual_permit_update_default,
        "sheet_name": "大小月卡+首充"
    }
}


TIER_TARGET_MAPPING = {
    "0氪": 0.5,           # 0氪看 50%
    "月卡": 0.6,          # 月卡看 60%
    "大小月卡": 0.7,       # 大小月卡看 70%
    "大小月卡+首充": 0.8   # 首充看 80%
}

COMPARISON_STRATEGIES_NAMES = ["120下池", "平铺60追up", "30+120"]

class GachaSimulator:
    def __init__(self, target_collection_percent, init_permit, update_permit, update_ars, annual_permit_update):
        self.attempt_5 = 0
        self.attempt_6_s = 0
        self.attempt_6_l = 0
        self.attempt_weapon = 0

        self.current_permits = init_permit
        self.update_permit = update_permit
        self.annual_permit_update = annual_permit_update
        self.special_permit = 0
        self.limited_permits = 0
        self.current_ars = 0
        self.update_ars = update_ars

        self.collection = [0] * 105
        self.weapons = [0] * 105
        self.target_collection_percent = target_collection_percent
        self.target = [0] * 105
        
        self.total_6_stars_obtained = 0

        for op_id in range(1, 105):
            if random.random() < self.target_collection_percent:
                self.target[op_id] = 1
            else:
                self.target[op_id] = 0

    # 下个池子，清空保底，计算工资
    def next_pool(self, current_pool_id):
        self.current_permits += self.update_permit
        self.current_ars += self.update_ars
        
        if current_pool_id % 52 == 1:
            self.current_permits += self.annual_permit_update

        base_limited = 10
        bonus_limited = 10 if self.special_permit == 1 else 0
        self.limited_permits = base_limited + bonus_limited
        
        self.special_permit = 0
        self.attempt_6_l = 0       
        self.attempt_weapon = 0    

    # 单抽
    def headhunt(self, operator_id):
        current_6_star_rate = 0.008
        if self.attempt_6_s >= 65: current_6_star_rate += (self.attempt_6_s - 64) * 0.05
        if self.attempt_6_s >= 79: current_6_star_rate = 1.0
        
        guaranteed_up = False
        if self.attempt_6_l >= 119:
            current_6_star_rate = 1.0
            guaranteed_up = True

        roll = random.random()
        result_id = -2 
        
        if roll < current_6_star_rate:
            self.total_6_stars_obtained += 1
            
            if guaranteed_up or random.random() < 0.5:
                result_id = operator_id
                self.attempt_6_l = 0; self.attempt_6_s = 0; self.attempt_5 = 0
            else:
                sub_roll = random.random()
                if operator_id == 1: result_id = 0 
                elif operator_id == 2:
                    if sub_roll < (1/7): result_id = 1 
                    else: result_id = 0
                else:
                    if sub_roll < (1/7): result_id = operator_id - 1
                    elif sub_roll < (2/7): result_id = operator_id - 2
                    else: result_id = 0 
                self.attempt_6_s = 0; self.attempt_5 = 0; self.attempt_6_l += 1
        elif self.attempt_5 >= 9 or roll < (current_6_star_rate + 0.08):
            result_id = -1
            self.attempt_5 = 0; self.attempt_6_s += 1; self.attempt_6_l += 1
        else:
            result_id = -2
            self.attempt_5 += 1; self.attempt_6_s += 1; self.attempt_6_l += 1

        if result_id >= 0: self.current_ars += 2000 
        elif result_id == -1: self.current_ars += 200
        else: self.current_ars += 20
        return result_id

    # 加急十连
    def urgent_headhunt(self, operator_id):
        results = []
        has_high_rarity_in_first_9 = False
        for i in range(10):
            rate_6 = 0.008; rate_5 = 0.08
            roll = random.random(); result_id = -2
            if roll < rate_6:
                self.total_6_stars_obtained += 1
                if random.random() < 0.5: result_id = operator_id
                else:
                    sub_roll = random.random()
                    if operator_id == 1: result_id = 0
                    elif operator_id == 2:
                        if sub_roll < (1/7): result_id = 1
                        else: result_id = 0
                    else:
                        if sub_roll < (1/7): result_id = operator_id - 1
                        elif sub_roll < (2/7): result_id = operator_id - 2
                        else: result_id = 0
            elif roll < (rate_6 + rate_5): result_id = -1
            else: result_id = -2
            
            if i == 9 and not has_high_rarity_in_first_9 and result_id == -2: result_id = -1
            if i < 9 and result_id > -2: has_high_rarity_in_first_9 = True
            results.append(result_id)
            
        for res in results:
            if res >= 0: self.current_ars += 2000
            elif res == -1: self.current_ars += 200
            else: self.current_ars += 20
        return results

    # 武器池
    def weapon_headhunt(self, operator_id):
        if self.current_ars < 1980: return False
        self.current_ars -= 1980
        got_weapon = False
        for _ in range(10):
            rate = 0.01
            if self.attempt_weapon >= 79: rate = 1.0
            if random.random() < rate:
                got_weapon = True
                self.weapons[operator_id] = 1 
                self.attempt_weapon = 0
            else: self.attempt_weapon += 1
        return got_weapon
    
    # 武器池策略：8次十连下池，直到出
    def solve_weapon_strategy(self, operator_id):
        if self.collection[operator_id] == 0: return
        if self.current_ars >= (1980 * 8):
            while self.weapons[operator_id] == 0:
                self.weapon_headhunt(operator_id)
                    
    # 生成模拟结果摘要
    def summary(self):
        total_6_stars = self.total_6_stars_obtained
        limited_6_stars = sum(self.collection[1:])
        unique_operators = 0
        for i in range(1, 105):
            if self.collection[i] > 0: unique_operators += 1
        collection_rate = unique_operators / 104
        total_targets = sum(self.target[1:])
        achieved_targets = 0
        for i in range(1, 105):
            if self.target[i] == 1 and self.collection[i] > 0: achieved_targets += 1
        target_success_rate = 0
        if total_targets > 0: target_success_rate = achieved_targets / total_targets
        total_weapons = sum(self.weapons[1:])
        weapon_match_count = 0
        for i in range(1, 105):
            if self.collection[i] > 0 and self.weapons[i] == 1: weapon_match_count += 1
        weapon_success_rate = 0
        if unique_operators > 0: weapon_success_rate = weapon_match_count / unique_operators
        return total_6_stars, limited_6_stars, collection_rate, target_success_rate, total_weapons, weapon_success_rate

# 120下池
def strategy_1(simulator, operator_id):
    initial_assets = simulator.current_permits + simulator.limited_permits
    can_go_all_in = (initial_assets >= 120)
    is_wanted = (simulator.target[operator_id] == 1)
    pool_pull_count = 0
    urgent_used = False 
    def _pull_once_with_result():
        nonlocal pool_pull_count, urgent_used
        if simulator.limited_permits > 0: simulator.limited_permits -= 1
        elif simulator.current_permits >= 1: simulator.current_permits -= 1
        else: return None 
        res = simulator.headhunt(operator_id)
        if res > 0: simulator.collection[res] += 1
        pool_pull_count += 1
        if pool_pull_count == 30 and not urgent_used:
            urgent_results = simulator.urgent_headhunt(operator_id)
            urgent_used = True
            for r in urgent_results:
                if r > 0: simulator.collection[r] += 1
        if pool_pull_count == 60: simulator.special_permit = 1
        return res
    while simulator.limited_permits > 0: _pull_once_with_result()
    if is_wanted and can_go_all_in and simulator.collection[operator_id] == 0:
        while simulator.collection[operator_id] == 0:
            success = _pull_once_with_result()
            if success is None: break 
    simulator.solve_weapon_strategy(operator_id)

# 平铺60追up
def strategy_3(simulator, operator_id):
    is_wanted = (simulator.target[operator_id] == 1)
    pool_pull_count = 0
    urgent_used = False 
    def _pull_once_with_result():
        nonlocal pool_pull_count, urgent_used
        if simulator.limited_permits > 0: simulator.limited_permits -= 1
        elif simulator.current_permits >= 1: simulator.current_permits -= 1
        else: return None 
        res = simulator.headhunt(operator_id)
        if res > 0: simulator.collection[res] += 1
        pool_pull_count += 1
        if pool_pull_count == 30 and not urgent_used:
            urgent_results = simulator.urgent_headhunt(operator_id)
            urgent_used = True
            for r in urgent_results:
                if r > 0: simulator.collection[r] += 1
        if pool_pull_count == 60: simulator.special_permit = 1
        return res
    while simulator.limited_permits > 0: _pull_once_with_result()
    if is_wanted and simulator.collection[operator_id] == 0:
        if pool_pull_count < 60:
            needed_for_ticket = 60 - pool_pull_count
            if simulator.current_permits >= needed_for_ticket:
                while pool_pull_count < 60 and simulator.collection[operator_id] == 0:
                    res = _pull_once_with_result()
                    if res is None: break 
        if simulator.collection[operator_id] == 0:
            needed_for_hard_pity = 120 - simulator.attempt_6_l
            if needed_for_hard_pity < 0: needed_for_hard_pity = 0
            if simulator.current_permits >= needed_for_hard_pity:
                while simulator.collection[operator_id] == 0:
                    res = _pull_once_with_result()
                    if res is None: break
    simulator.solve_weapon_strategy(operator_id)

# 平铺60
def strategy_3_1(simulator, operator_id):
    pool_pull_count = 0
    urgent_used = False 
    def _pull_once_with_result():
        nonlocal pool_pull_count, urgent_used
        if simulator.limited_permits > 0: simulator.limited_permits -= 1
        elif simulator.current_permits >= 1: simulator.current_permits -= 1
        else: return None 
        res = simulator.headhunt(operator_id)
        if res > 0: simulator.collection[res] += 1
        pool_pull_count += 1
        if pool_pull_count == 30 and not urgent_used:
            urgent_results = simulator.urgent_headhunt(operator_id)
            urgent_used = True
            for r in urgent_results:
                if r > 0: simulator.collection[r] += 1
        if pool_pull_count == 60: simulator.special_permit = 1
        return res
    while pool_pull_count < 60:
        res = _pull_once_with_result()
        if res is None: break
        if simulator.collection[operator_id] > 0:
            if pool_pull_count >= 50: pass
            else: break
    simulator.solve_weapon_strategy(operator_id)

# 30+120
def strategy_4(simulator, operator_id):
    is_wanted = (simulator.target[operator_id] == 1)
    pool_pull_count = 0
    urgent_used = False 
    def _pull_once_with_result():
        nonlocal pool_pull_count, urgent_used
        if simulator.limited_permits > 0: simulator.limited_permits -= 1
        elif simulator.current_permits >= 1: simulator.current_permits -= 1
        else: return None 
        res = simulator.headhunt(operator_id)
        if res > 0: simulator.collection[res] += 1
        pool_pull_count += 1
        if pool_pull_count == 30 and not urgent_used:
            urgent_results = simulator.urgent_headhunt(operator_id)
            urgent_used = True
            for r in urgent_results:
                if r > 0: simulator.collection[r] += 1
        if pool_pull_count == 60: simulator.special_permit = 1
        return res
    while simulator.limited_permits > 0:
        res = _pull_once_with_result()
    needed_for_current = 120 - simulator.attempt_6_l
    if needed_for_current < 0: needed_for_current = 0
    if is_wanted and simulator.collection[operator_id] == 0:
        if simulator.current_permits >= needed_for_current:
            while simulator.collection[operator_id] == 0:
                res = _pull_once_with_result()
                if res is None: break
    if simulator.collection[operator_id] > 0:
        if operator_id < 104 and 47 <= pool_pull_count < 60:
            if simulator.target[operator_id + 1] == 1:
                cost_to_top_up = 60 - pool_pull_count
                cost_next_guarantee = 120
                remaining_stock = simulator.current_permits - cost_to_top_up
                future_assets = remaining_stock + simulator.update_permit + 10 + 10
                if future_assets >= cost_next_guarantee:
                    while pool_pull_count < 60:
                        res = _pull_once_with_result()
                        if res is None: break
    if simulator.collection[operator_id] == 0:
        if operator_id < 104:
            next_is_wanted = (simulator.target[operator_id + 1] == 1)
            if next_is_wanted and simulator.attempt_6_s < 30:
                cost_to_pad = 30 - simulator.attempt_6_s
                cost_next_guarantee = 120 
                total_needed = cost_to_pad + cost_next_guarantee
                future_assets = simulator.current_permits + simulator.update_permit + 10
                if future_assets >= total_needed:
                    while simulator.attempt_6_s < 30:
                        res = _pull_once_with_result()
                        if res is None: break 
                        if res >= 0: break 
                        if simulator.collection[operator_id] > 0: break 
    simulator.solve_weapon_strategy(operator_id)

def format_stats(data_list, is_percent=False):
    # 格式: (Mean - 2SD) - Mean - (Mean + 2SD)
    if not data_list:
        return "N/A"
    mu = statistics.mean(data_list)
    sigma = statistics.stdev(data_list)
    lower_bound = mu - 2 * sigma
    upper_bound = mu + 2 * sigma
    if is_percent:
        if upper_bound > 1.0: upper_bound = 1.0
        if lower_bound < 0.0: lower_bound = 0.0
        return f"{lower_bound:.2%}-{mu:.2%}-{upper_bound:.2%}"
    else:
        if lower_bound < 0: lower_bound = 0
        return f"{lower_bound:.2f}-{mu:.2f}-{upper_bound:.2f}"

def get_mean(data_list):
    if not data_list: return 0
    return statistics.mean(data_list)

if __name__ == "__main__":
    TARGET_RANGE = [0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0]
    
    TIER_ORDER = [
        "TIER_0 (0氪)",
        "TIER_1 (月卡)",
        "TIER_2 (大月卡)",
        "TIER_3 (首充双倍)"
    ]
    
    STRATEGIES = [
        (strategy_1, "120下池"),
        (strategy_3, "平铺60追up"),
        (strategy_3_1, "平铺60"), 
        (strategy_4, "30+120")
    ]
    
    HEADERS = [
        "Target %", "总计6星 (包括常驻)", "限定6星",           
        "图鉴收集率", "目标图鉴达成率", "总计专武数", "角色专武百分比"       
    ]
    
    
    OUTPUT_FILE = "final_comparison.xlsx"
    
    start_total = time.time()
    workbook = xlsxwriter.Workbook(OUTPUT_FILE)
    
    header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D7E4BC', 'border': 1})
    center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
    section_format = workbook.add_format({'bold': True, 'bg_color': '#FFEB9C', 'border': 1})
    
    comp_header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFE699', 'border': 1}) # 浅橙色表头
    comp_title_format = workbook.add_format({'bold': True, 'bg_color': '#FFFF00', 'border': 1}) # 亮黄色标题
    highlight_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFFF00', 'border': 1}) # 最优解黄色
    normal_comp_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1}) # 普通解

    for tier_name in TIER_ORDER:
        cfg = CONFIGS[tier_name]
        sheet_name = cfg["sheet_name"] 
        target_compare_percent = TIER_TARGET_MAPPING[sheet_name] 

        print(f"\n>>> 正在生成 Sheet: {sheet_name} (横向对比目标: {target_compare_percent:.0%}) ...")
        
        worksheet = workbook.add_worksheet(sheet_name)
        worksheet.set_column('A:A', 20) 
        worksheet.set_column('B:G', 25) 
        
        current_row = 0
        
        sheet_stats_buffer = {}

        for strat_func, strat_name in STRATEGIES:
            print(f"  - Running {strat_name}...")
            
            if strat_name not in sheet_stats_buffer:
                sheet_stats_buffer[strat_name] = {}

            worksheet.merge_range(current_row, 0, current_row, 6, strat_name, section_format)
            current_row += 1
            worksheet.write_row(current_row, 0, HEADERS, header_format)
            current_row += 1
            
            if strat_func == strategy_3_1:
                current_target_range = [1.0]
            else:
                current_target_range = TARGET_RANGE
            
            for target_p in current_target_range:
                d_tot6, d_lim6, d_col, d_tar, d_twep, d_wrate = [], [], [], [], [], []
                
                for _ in range(SIMULATION_COUNT):
                    sim = GachaSimulator(target_p, cfg['init'], cfg['update'], cfg['ars'], cfg['annual'])
                    for pool_id in range(1, 105):
                        sim.next_pool(pool_id)
                        strat_func(sim, pool_id)
                    
                    t6, l6, cr, tr, tw, wr = sim.summary()
                    d_tot6.append(t6); d_lim6.append(l6); d_col.append(cr)
                    d_tar.append(tr); d_twep.append(tw); d_wrate.append(wr)
                
                # 写入主表 (Mean +- 2SD)
                row_data = [
                    f"{target_p:.0%}",
                    format_stats(d_tot6, False), format_stats(d_lim6, False),
                    format_stats(d_col, True), format_stats(d_tar, True),
                    format_stats(d_twep, False), format_stats(d_wrate, True)
                ]
                worksheet.write_row(current_row, 0, row_data, center_format)
                current_row += 1
                
                sheet_stats_buffer[strat_name][target_p] = [
                    get_mean(d_tot6), get_mean(d_lim6), 
                    get_mean(d_col), get_mean(d_tar),
                    get_mean(d_twep), get_mean(d_wrate)
                ]
            
            current_row += 1 

        current_row += 1 
        comp_title = f"{target_compare_percent:.0%}目标下，3种策略的横向对比"
        
        worksheet.merge_range(current_row, 0, current_row, 6, comp_title, comp_title_format)
        current_row += 1
        
        comp_headers = ["策略"] + HEADERS[1:] 
        worksheet.write_row(current_row, 0, comp_headers, comp_header_format)
        current_row += 1
        
        compare_strats = ["120下池", "平铺60追up", "30+120"]
        
        col_max_values = []
        
        for col_idx in range(6): 
            vals = []
            for s_name in compare_strats:
                if s_name in sheet_stats_buffer and target_compare_percent in sheet_stats_buffer[s_name]:
                    vals.append(sheet_stats_buffer[s_name][target_compare_percent][col_idx])
                else:
                    vals.append(0)
            col_max_values.append(max(vals) if vals else 0)

        for s_name in compare_strats:
            worksheet.write(current_row, 0, s_name, normal_comp_format)
            
            if s_name in sheet_stats_buffer and target_compare_percent in sheet_stats_buffer[s_name]:
                means = sheet_stats_buffer[s_name][target_compare_percent]
                
                for col_idx in range(6):
                    val = means[col_idx]
                    max_val = col_max_values[col_idx]
                    
                    is_pct_col = (col_idx in [2, 3, 5])
                    
                    if is_pct_col:
                        val_display = f"{val:.2%}"
                    else:
                        val_display = f"{val:.2f}"
                        
                    # 计算差距
                    diff = 0
                    if max_val > 0:
                        diff = (val - max_val) / max_val
                    
                    if abs(diff) < 0.0001: 
                        final_text = f"{val_display}(0%)"
                        cell_fmt = highlight_format
                    else:
                        final_text = f"{val_display}({diff:.2%})"
                        cell_fmt = normal_comp_format
                    
                    worksheet.write(current_row, col_idx + 1, final_text, cell_fmt)
            else:
                worksheet.write(current_row, 1, "No Data", normal_comp_format)

            current_row += 1

    workbook.close()
    print("\n" + "=" * 40)
    print(f"模拟结束！总耗时: {time.time() - start_total:.1f}s")
    print(f"请查看文件: {OUTPUT_FILE}")
