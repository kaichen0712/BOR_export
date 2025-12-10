# scheduler.py
# ---------------------------------------------------------
# 將 parser 產出的事件映射，展開成「每天」的班別表
# - 先依 weekend.json / holidays.json 預填「例假」「國定假」
# - 事件（休假、班別）會覆蓋預設
# ---------------------------------------------------------

import json
from datetime import datetime
from pathlib import Path
from typing import Dict

BASE_DIR = Path(__file__).resolve().parent


def _load_dates_from_json(filename: str):
    path = BASE_DIR / filename
    try:
        with open(path, "r", encoding="utf-8") as f:
            return set(json.load(f))
    except FileNotFoundError:
        return set()


def _all_dates_in_month(year: int, month: int):
    days = []
    day = 1
    while True:
        try:
            dt = datetime(year, month, day)
        except ValueError:
            break
        days.append(dt.strftime("%Y-%m-%d"))
        day += 1
    return days


def expand_to_daily(schedule_dict: Dict[str, Dict[str, str]], base_year: int, target_month: int, identity_map: Dict[str, str] = None):
    """
    將人名 -> {date: shift} 的事件表，展開成全月每日表。
    
    根據身分別設定不同的假日標籤：
    
    契約人員：
    - 禮拜六 → 休息日
    - 禮拜日 → 例假
    - 國定假 → 國定假
    - 公休/休假 → 特別休假
    
    公職人員：
    - 禮拜六、禮拜日 → 例假
    - 國定假 → 國定假（會被公休覆蓋為休假）
    - 公休/休假 → 休假
    """
    if identity_map is None:
        identity_map = {}
    
    weekend_dates = _load_dates_from_json("weekend.json")
    holiday_dates = _load_dates_from_json("holidays.json")
    all_dates = _all_dates_in_month(base_year, target_month)

    daily: Dict[str, Dict[str, str]] = {}

    for name, events in schedule_dict.items():
        daily[name] = {}
        
        # 取得身分別（預設為公職）
        identity = identity_map.get(name, '')  # 找不到就留空

        # 預填假日（根據身分別）
        for date_key in all_dates:
            label = None
            dt = datetime.strptime(date_key, "%Y-%m-%d")
            weekday = dt.weekday()  # 0=Mon, 5=Sat, 6=Sun
            
            if date_key in holiday_dates:
                # 國定假日：兩種身分都是「國定假」
                label = "國定假"
            elif date_key in weekend_dates:
                if identity == '契約':
                    # 契約人員：禮拜六=休息日，禮拜日=例假
                    if weekday == 5:  # Saturday
                        label = "休息日"
                    elif weekday == 6:  # Sunday
                        label = "例假"
                else:
                    # 公職人員：禮拜六、禮拜日=例假
                    label = "例假"

            if label:
                daily[name][date_key] = label

        # 事件覆蓋（根據身分別轉換休假標籤）
        for date_key, shift in events.items():
            if shift == "休假":
                # 檢查是否為週末（週六=5, 週日=6）
                dt = datetime.strptime(date_key, "%Y-%m-%d")
                weekday = dt.weekday()
                
                if weekday in [5, 6]:
                    # 週六日不覆蓋，保留原本的值（例假/休息日）
                    continue
                
                # 公職人員：國定假日不覆蓋，保留「國定假」
                if identity == '公職' and date_key in holiday_dates:
                    continue
                
                if identity == '契約':
                    # 契約人員的公休（週一~五）→ 特別休假
                    daily[name][date_key] = "特別休假"
                else:
                    # 公職人員的公休（週一~五，非國定假）→ 休假
                    daily[name][date_key] = "休假"
            else:
                daily[name][date_key] = shift

    return daily


if __name__ == "__main__":
    # 簡易測試：輸入一個事件，觀察展開
    sample = {"王小明": {"2026-02-05": "7~3"}}
    expanded = expand_to_daily(sample, 2026, 2)
    print(expanded["王小明"])