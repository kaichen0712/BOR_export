# parser_v2.py
# ---------------------------------------------------------
# 解析 11502test.xlsx（主 & 副值資料）
# 修正版本 v2：
# - 新增「備註」欄位支援
# - 判斷 X月換心、X月大/小P1、X月大/小P2 的月份是否符合 target_month
# - 只有月份符合才執行對應排班規則
# ---------------------------------------------------------

import pandas as pd
import re
from datetime import datetime, timedelta
import json
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent


def load_holidays(json_path=None):
    if json_path is None:
        json_path = BASE_DIR / "holidays.json"
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            holidays = json.load(f)
        return set(holidays)
    except:
        return set()


def load_identity_map(filepath, sheet_name='身分'):
    """
    讀取身分表，返回 {姓名: 身分} 的對應字典
    身分為 '公職' 或 '契約'
    """
    try:
        df = pd.read_excel(filepath, sheet_name=sheet_name, header=1)
        identity_map = {}
        for _, row in df.iterrows():
            # 取得姓名（移除空白）
            raw_name = str(row.get('姓名', '')).strip()
            name = re.sub(r'[\s\u3000]+', '', raw_name)
            identity = str(row.get('身分', '')).strip()
            if name and identity in ['公職', '契約']:
                identity_map[name] = identity
        print(f"[DEBUG] 載入身分表：公職 {sum(1 for v in identity_map.values() if v == '公職')} 人，契約 {sum(1 for v in identity_map.values() if v == '契約')} 人")
        return identity_map
    except Exception as e:
        print(f"[WARNING] 無法讀取身分表: {e}")
        return {}


def load_identity_map_from_bytes(file_bytes, sheet_name='身分'):
    """
    從記憶體中的 bytes 讀取身分表
    """
    import io
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=1)
        identity_map = {}
        for _, row in df.iterrows():
            raw_name = str(row.get('姓名', '')).strip()
            name = re.sub(r'[\s\u3000]+', '', raw_name)
            identity = str(row.get('身分', '')).strip()
            if name and identity in ['公職', '契約']:
                identity_map[name] = identity
        print(f"[DEBUG] 載入身分表：公職 {sum(1 for v in identity_map.values() if v == '公職')} 人，契約 {sum(1 for v in identity_map.values() if v == '契約')} 人")
        return identity_map
    except Exception as e:
        print(f"[WARNING] 無法讀取身分表: {e}")
        return {}


def clean_name(name_str):
    """移除姓名前後的英文、數字、空白、星號"""
    if not name_str:
        return ""
    name_str = str(name_str)
    name_str = name_str.translate(
        str.maketrans(
            {chr(full): chr(full - 0xFEE0) for full in range(ord("！"), ord("～") + 1)}
        )
    )
    name_str = name_str.replace("\u3000", " ")
    name_str = re.sub(r'^[A-Za-z0-9\s＊\*]+', "", name_str)
    name_str = re.sub(r'[A-Za-z0-9\s＊\*]+$', "", name_str)
    name_str = re.sub(r'[\s\u3000]+', " ", name_str)
    return name_str.strip()


def is_holiday(date_obj, holidays):
    """判斷是否假日（週六、週日、國定假日）"""
    if date_obj.weekday() >= 5:
        return True
    date_str = date_obj.strftime("%Y-%m-%d")
    return date_str in holidays


def parse_date_str(date_str, base_year, target_month):
    """
    將字串日期（如 1/15 或 *1/15）轉成 datetime 物件
    """
    if "/" not in str(date_str):
        return None
    
    # 移除開頭的 * 符號和其他非數字字元
    cleaned = re.sub(r'^[^\d]+', '', str(date_str).strip())
    if "/" not in cleaned:
        return None
    
    try:
        parts = cleaned.split("/")
        m = int(parts[0])
        d = int(parts[1])
    except:
        return None
    
    year = base_year
    # 如果月份比目標月份大很多（例如 12 月而目標是 2 月），可能是前一年
    if m > target_month + 6:
        year = base_year - 1
    
    try:
        return datetime(year, m, d)
    except ValueError:
        return None


def parse_multi_dates(cell, base_year, target_month):
    """解析 '1/15.1/16.1/17' 或 '*1/15.1/16.1/17' 多日期格式"""
    # 先移除開頭的 * 符號
    cleaned = str(cell).strip().lstrip('*＊')
    dates = []
    for part in cleaned.split("."):
        dt = parse_date_str(part.strip(), base_year, target_month)
        if dt:
            dates.append(dt)
    return dates


def parse_range(cell, base_year, target_month):
    """解析日期區間 1/12–1/18 或 1/12-1/18 或 *1/12-1/16"""
    # 先移除開頭的 * 符號
    cleaned = str(cell).strip().lstrip('*＊')
    m = re.search(r"(\d+/\d+)[–\-](\d+/\d+)", cleaned)
    if not m:
        return []
    left, right = m.groups()
    start = parse_date_str(left.strip(), base_year, target_month)
    end = parse_date_str(right.strip(), base_year, target_month)
    if not start or not end:
        return []
    
    if end < start:
        end = datetime(end.year + 1, end.month, end.day)
    
    delta = (end - start).days
    return [start + timedelta(days=i) for i in range(delta + 1)]


def parse_holiday_cell(cell, base_year, target_month):
    """解析假日欄位格式：'1/4白班', '1/4小夜', '1/4大夜'"""
    m = re.match(r"(\d+/\d+)(白班|小夜|大夜)", cell)
    if not m:
        return None, None
    date_str, type_str = m.groups()
    dt = parse_date_str(date_str, base_year, target_month)
    if not dt:
        return None, None
    if type_str == "白班":
        shift = "7~3"
    elif type_str == "小夜":
        shift = "3~11"
    elif type_str == "大夜":
        shift = "23~7"
    else:
        shift = ""
    return dt, shift


def parse_compensate(cell, base_year, target_month):
    """
    補休格式：「補1/1 (原1/17白班)」或「公休(原1/31小夜)」
    返回：休假日期（原班那天）
    """
    # 先嘗試找「原X/X」
    m_orig = re.search(r"原(\d+/\d+)", cell)
    if m_orig:
        date_str = m_orig.group(1)
        return parse_date_str(date_str, base_year, target_month)
    
    # 如果沒有「原」，就用「補」後面的日期
    m_comp = re.match(r"補(\d+/\d+)", cell)
    if m_comp:
        date_str = m_comp.group(1)
        return parse_date_str(date_str, base_year, target_month)
    
    return None


def extract_month_from_rule(text):
    """
    從規則文字中提取月份
    例如：「2月換心」→ 2, 「1月大P1」→ 1, 「12月小P2」→ 12
    返回：月份數字 或 None
    """
    # 匹配 "X月" 格式
    m = re.match(r"(\d+)月", text)
    if m:
        return int(m.group(1))
    return None


def parse_monthly_rules(text, target_month, base_year, holidays, schedule_dict_name, identity='公職'):
    """
    解析月份規則（換心、P1、P2）
    只有當規則中的月份 == target_month 時才執行
    
    參數：
    - text: 規則文字，如「2月換心」、「2月大P1」
    - target_month: 目標月份
    - base_year: 基準年份
    - holidays: 國定假日集合
    - schedule_dict_name: 該員工的排班字典（會被修改）
    - identity: 身分別（'公職' 或 '契約'）
    
    返回：是否有處理到規則
    """
    processed = False
    
    # 換心規則：X月換心
    if "換心" in text:
        rule_month = extract_month_from_rule(text)
        if rule_month == target_month:
            for day in range(1, 32):
                try:
                    dt = datetime(base_year, target_month, day)
                except:
                    break
                key = dt.strftime("%Y-%m-%d")
                if is_holiday(dt, holidays):
                    schedule_dict_name[key] = "休假"
                else:
                    schedule_dict_name[key] = "7~3"
            processed = True
            print(f"    [INFO] 套用 {text} 規則")
    
    # P1 規則：X月大P1 或 X月小P1
    # 週日=例假、週一=7~3、週二=7~3
    # 但國定假日不覆蓋，維持「國定假」
    elif "P1" in text:
        rule_month = extract_month_from_rule(text)
        if rule_month == target_month:
            for day in range(1, 32):
                try:
                    dt = datetime(base_year, target_month, day)
                except:
                    break
                key = dt.strftime("%Y-%m-%d")
                wd = dt.weekday()  # 0=Mon, 6=Sun
                
                # 國定假日跳過，不覆蓋
                if key in holidays:
                    continue
                
                if wd == 6:  # 週日
                    schedule_dict_name[key] = "例假"
                elif wd in [0, 1]:  # 週一、週二
                    schedule_dict_name[key] = "7~3"
            processed = True
            print(f"    [INFO] 套用 {text} 規則")
    
    # P2 規則：X月大P2 或 X月小P2
    # 週三=7~3、週四=7~3、週五=7~3
    # 週六：公職=例假、契約=休息日
    # 但國定假日不覆蓋，維持「國定假」
    elif "P2" in text:
        rule_month = extract_month_from_rule(text)
        if rule_month == target_month:
            for day in range(1, 32):
                try:
                    dt = datetime(base_year, target_month, day)
                except:
                    break
                key = dt.strftime("%Y-%m-%d")
                wd = dt.weekday()
                
                # 國定假日跳過，不覆蓋
                if key in holidays:
                    continue
                
                if wd == 5:  # 週六
                    if identity == '契約':
                        schedule_dict_name[key] = "休息日"
                    else:
                        schedule_dict_name[key] = "例假"
                elif wd in [2, 3, 4]:  # 週三、週四、週五
                    schedule_dict_name[key] = "7~3"
            processed = True
            print(f"    [INFO] 套用 {text} 規則（身分：{identity}）")
    
    return processed


def _find_sheet_names(excel_file):
    """
    根據頁籤名稱找出主值、副值、身分頁籤
    頁籤名稱格式例如：11502(主)、11502(副)、身分
    """
    xl = pd.ExcelFile(excel_file)
    sheet_names = xl.sheet_names
    
    main_sheet = None
    sub_sheet = None
    identity_sheet = None
    
    for name in sheet_names:
        name_str = str(name)
        if '主' in name_str:
            main_sheet = name
        elif '副' in name_str:
            sub_sheet = name
        elif '身分' in name_str or '身份' in name_str:
            identity_sheet = name
    
    # 如果找不到，使用預設索引
    if main_sheet is None and len(sheet_names) >= 1:
        main_sheet = sheet_names[0]
        print(f"[WARNING] 找不到主值頁籤，使用第一個頁籤: {main_sheet}")
    if sub_sheet is None and len(sheet_names) >= 2:
        sub_sheet = sheet_names[1]
        print(f"[WARNING] 找不到副值頁籤，使用第二個頁籤: {sub_sheet}")
    if identity_sheet is None:
        # 嘗試找「身分」頁籤
        for name in sheet_names:
            if '身' in str(name):
                identity_sheet = name
                break
        if identity_sheet is None and len(sheet_names) >= 3:
            identity_sheet = sheet_names[2]
            print(f"[WARNING] 找不到身分頁籤，使用第三個頁籤: {identity_sheet}")
    
    print(f"[DEBUG] 頁籤對應 - 主值: {main_sheet}, 副值: {sub_sheet}, 身分: {identity_sheet}")
    return main_sheet, sub_sheet, identity_sheet


def parse_source_excel_from_bytes(file_bytes, base_year, target_month):
    """從記憶體中的 bytes 解析 Excel"""
    import io
    holidays = load_holidays()
    
    file_io = io.BytesIO(file_bytes)
    main_sheet, sub_sheet, identity_sheet = _find_sheet_names(file_io)
    
    file_io.seek(0)
    identity_map = load_identity_map_from_bytes(file_bytes, sheet_name=identity_sheet)
    
    file_io.seek(0)
    df_main = pd.read_excel(file_io, sheet_name=main_sheet, header=1)
    file_io.seek(0)
    df_sub = pd.read_excel(file_io, sheet_name=sub_sheet, header=1)
    
    return _parse_dataframes(df_main, df_sub, holidays, identity_map, base_year, target_month)


def parse_source_excel(filepath, base_year, target_month):
    holidays = load_holidays()
    
    main_sheet, sub_sheet, identity_sheet = _find_sheet_names(filepath)
    identity_map = load_identity_map(filepath, sheet_name=identity_sheet)

    df_main = pd.read_excel(filepath, sheet_name=main_sheet, header=1)
    df_sub = pd.read_excel(filepath, sheet_name=sub_sheet, header=1)
    
    return _parse_dataframes(df_main, df_sub, holidays, identity_map, base_year, target_month)


def _parse_dataframes(df_main, df_sub, holidays, identity_map, base_year, target_month):
    """解析主值和副值的 DataFrame（共用邏輯）"""

    # 統一姓名欄位名稱
    if '主值' in df_main.columns:
        df_main = df_main.rename(columns={'主值': '姓名'})
    if '副值' in df_sub.columns:
        df_sub = df_sub.rename(columns={'副值': '姓名'})

    def _normalize_cols(df):
        """合併所有可能有多個的欄位"""
        
        def merge_columns(df, prefix, target_name, separator):
            """合併以 prefix 開頭的所有欄位"""
            cols = [c for c in df.columns if str(c).startswith(prefix)]
            if cols:
                df[target_name] = df[cols].apply(
                    lambda row: separator.join(
                        [
                            str(v).strip()
                            for v in row
                            if pd.notna(v) and str(v).strip() not in ["", "nan"]
                        ]
                    ),
                    axis=1,
                )
                drop_cols = [c for c in cols if c != target_name]
                if drop_cols:
                    df = df.drop(columns=drop_cols)
            return df
        
        # 合併假日欄（假日、假日.1、假日.2...）
        df = merge_columns(df, "假日", "假日", "、")
        
        # 合併大夜欄（大夜、大夜.1、大夜.2...）
        df = merge_columns(df, "大夜", "大夜", ".")
        
        # 合併小夜週欄（小夜週、小夜週.1、小夜...）
        df = merge_columns(df, "小夜", "小夜週", ".")
        
        # 合併公休欄（公休、公休.1、公休.2...）
        df = merge_columns(df, "公休", "公休", "、")
        
        # 合併備註欄（備註、備註.1、備註.2...）
        df = merge_columns(df, "備註", "備註", "、")
        
        return df

    df_main = _normalize_cols(df_main)
    df_sub = _normalize_cols(df_sub)
    
    # 收集主值表和副值表的人員清單（用於最後比對）
    main_names = set()
    for _, row in df_main.iterrows():
        name = clean_name(row.iloc[0])
        if name and name != "nan":
            main_names.add(name)
    
    sub_names = set()
    for _, row in df_sub.iterrows():
        name = clean_name(row.iloc[0])
        if name and name != "nan":
            sub_names.add(name)
    
    df_all = pd.concat([df_main, df_sub], ignore_index=True)

    schedule_dict = {}
    monthly_rules_applied = {}  # 記錄套用月份規則的人員 {name: [rule1, rule2...]}

    for idx, row in df_all.iterrows():
        name = clean_name(row.iloc[0])
        if not name or name == "nan":
            continue
        
        # 過濾掉明顯不是人名的資料
        # 1. 名字太長（超過6個字，正常中文人名2-4字）
        # 2. 以數字開頭（可能是公告編號如「1.為安排...」）
        # 3. 以「.」開頭（公告文字）
        # 4. 包含標點符號
        # 5. 名字太短（少於2個字）
        if len(name) > 6:
            continue
        if len(name) < 2:
            continue
        if re.match(r'^[\d\.]', name):  # 以數字或「.」開頭
            continue
        if re.search(r'[。，、；：「」『』【】（）\(\)\.,;:\[\]]', name):
            continue

        schedule_dict[name] = {}
        identity = identity_map.get(name, '')  # 取得身分別，找不到就留空
        print(f"[DEBUG] 處理員工: {name}（{identity}）")

        # =============================================
        # 1. 先處理「備註」欄位（如果有的話）
        # =============================================
        note_col = str(row.get("備註", "")).strip()
        if note_col and note_col != "nan":
            parts = note_col.replace(" ", "").replace("\n", "、").split("、")
            for p in parts:
                if not p:
                    continue
                if parse_monthly_rules(p, target_month, base_year, holidays, schedule_dict[name], identity):
                    # 記錄套用的規則
                    if name not in monthly_rules_applied:
                        monthly_rules_applied[name] = []
                    monthly_rules_applied[name].append(p)

        # =============================================
        # 2. 公休欄（換心/P1/P2 規則 + 日期休假）
        # =============================================
        rest_col = str(row.get("公休", "")).strip()
        if rest_col and rest_col != "nan":
            parts = rest_col.replace(" ", "").replace("\n", "、").split("、")
            for p in parts:
                if not p:
                    continue
                
                # 跳過備註類文字（如「*滿55歲跳大夜」）
                if p.startswith("*"):
                    continue
                
                # 月份規則（換心、P1、P2）
                if "換心" in p or "P1" in p or "P2" in p:
                    if parse_monthly_rules(p, target_month, base_year, holidays, schedule_dict[name], identity):
                        # 記錄套用的規則
                        if name not in monthly_rules_applied:
                            monthly_rules_applied[name] = []
                        monthly_rules_applied[name].append(p)
                    continue

                # 補休格式：「補1/1 (原1/17白班)」或「公休(原1/31小夜)」
                if "補" in p or "原" in p:
                    dt = parse_compensate(p, base_year, target_month)
                    if dt:
                        key = dt.strftime("%Y-%m-%d")
                        schedule_dict[name][key] = "休假"
                        print(f"    [INFO] 補休: {key} -> 休假")
                    continue

                # 日期區間
                if "–" in p or "-" in p:
                    dates = parse_range(p, base_year, target_month)
                    for dt in dates:
                        key = dt.strftime("%Y-%m-%d")
                        schedule_dict[name][key] = "休假"
                    if dates:
                        print(f"    [INFO] 公休區間: {len(dates)} 天")
                    continue

                # 單一日期
                if "/" in p:
                    dt = parse_date_str(p, base_year, target_month)
                    if dt:
                        key = dt.strftime("%Y-%m-%d")
                        schedule_dict[name][key] = "休假"

        # =============================================
        # 3. 大夜欄
        # =============================================
        night_col = str(row.get("大夜", "")).strip()
        if night_col and night_col != "nan":
            if "." in night_col:
                dates = parse_multi_dates(night_col, base_year, target_month)
            elif "–" in night_col or "-" in night_col:
                dates = parse_range(night_col, base_year, target_month)
            else:
                dates = [parse_date_str(night_col, base_year, target_month)]
            for dt in dates:
                if dt:
                    key = dt.strftime("%Y-%m-%d")
                    schedule_dict[name][key] = "23~7"

        # =============================================
        # 4. 小夜週欄
        # =============================================
        eve_col = str(row.get("小夜週", "")).strip()
        if eve_col and eve_col != "nan":
            if "–" in eve_col or "-" in eve_col:
                dates = parse_range(eve_col, base_year, target_month)
            elif "." in eve_col:
                dates = parse_multi_dates(eve_col, base_year, target_month)
            else:
                dates = [parse_date_str(eve_col, base_year, target_month)]
            for dt in dates:
                if dt:
                    key = dt.strftime("%Y-%m-%d")
                    schedule_dict[name][key] = "3~11"

        # =============================================
        # 5. 假日欄（日期+班別，跳過月份規則）
        # =============================================
        holi_col = str(row.get("假日", "")).strip()
        if holi_col and holi_col != "nan":
            items = holi_col.replace(" ", "").split("、")
            for p in items:
                # 跳過月份規則（如「2月大P1」、「2月換心」）
                if re.match(r"\d+月", p):
                    continue
                # 跳過「X月補」格式
                if "月補" in p:
                    continue
                    
                dt, shift = parse_holiday_cell(p, base_year, target_month)
                if dt and shift:
                    key = dt.strftime("%Y-%m-%d")
                    schedule_dict[name][key] = shift

    # =============================================
    # 6. 為身分表中所有人員建立空排班記錄（如果尚未存在）
    #    這樣 scheduler 才能為他們設定正確的假日標籤
    # =============================================
    for name in identity_map.keys():
        if name not in schedule_dict:
            schedule_dict[name] = {}
            print(f"[DEBUG] 新增身分表人員（無排班資料）: {name}")

    total_entries = sum(len(v) for v in schedule_dict.values())
    print(f"\n[DEBUG] 解析完成：人數={len(schedule_dict)}，班別筆數={total_entries}")
    print(f"[DEBUG] 套用月份規則人員：{list(monthly_rules_applied.keys())}")
    
    # =============================================
    # 7. 輸出身分表中有但主值/副值表都沒有比對到的人員清單
    # =============================================
    all_shift_names = main_names | sub_names  # 主值表 + 副值表的所有人員
    identity_only = set(identity_map.keys()) - all_shift_names  # 身分表有但班表沒有的人
    shift_only = all_shift_names - set(identity_map.keys())  # 班表有但身分表沒有的人
    
    if identity_only:
        print(f"\n[WARNING] 身分表中有，但主值/副值表都沒有比對到的人員（共 {len(identity_only)} 人）：")
        for name in sorted(identity_only):
            print(f"    - {name}（{identity_map[name]}）")
    else:
        print(f"\n[INFO] 身分表人員皆有對應到主值或副值表")
    
    if shift_only:
        print(f"\n[WARNING] 主值/副值表中有，但身分表沒有比對到的人員（共 {len(shift_only)} 人）：")
        for name in sorted(shift_only):
            in_main = "主值" if name in main_names else ""
            in_sub = "副值" if name in sub_names else ""
            source = "、".join(filter(None, [in_main, in_sub]))
            print(f"    - {name}（來源：{source}）")
    
    return schedule_dict, monthly_rules_applied


if __name__ == "__main__":
    result, monthly_rules = parse_source_excel("11502test.xlsx", 2026, 2)
    print("\n=== 前 5 位員工的排班 ===")
    for k, v in list(result.items())[:5]:
        print(f"{k}: {dict(list(v.items())[:5])}")
    print(f"\n=== 套用月份規則人員 ===")
    for k, v in monthly_rules.items():
        print(f"{k}: {v}")