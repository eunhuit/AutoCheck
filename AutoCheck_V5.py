import tkinter as tk
import tkinter.font as tkFont
import tkinter.simpledialog
import tkinter.messagebox 
import json
import sys
import os
from datetime import datetime, timedelta
import gspread
from gspread.utils import rowcol_to_a1

# 디버깅
DEBUG_MODE = False  # 활성화: True, 비활성화: False
DEBUG_DATE = datetime(2025, 3, 26, 10, 0, 0)  # 디버깅 모드일 때 사용할 날짜

def get_now():
    return DEBUG_DATE if DEBUG_MODE else datetime.now()

# 설정 저장 파일
CONFIG_FILE = "config.json"

# 기본 설정 불러오기/저장 함수
def load_config():
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            config = json.load(f)
    except FileNotFoundError:
        config = {
            "BASE_ROW": 25,
            "DEPARTMENT_NAME": "IT네트워크시스템",
            "USER_NAME": "",  # 기본 사용자 이름은 빈 문자열
            "WEEK_NUMBER": "1째주",  # 드롭다운에서 선택할 기본값
            "MANUAL_MONTH": "",  # 빈 문자열이면 자동으로 현재 월 사용, 아니면 수동 선택한 월 사용 (예: "4월")
            # 커스텀 디데이 설정 (없으면 None)
            "CUSTOM_DDAY": {"label": "", "date": ""}
        }
    if "WEEK_NUMBER" not in config:
        config["WEEK_NUMBER"] = "1째주"
    if "CUSTOM_DDAY" not in config:
        config["CUSTOM_DDAY"] = {"label": "", "date": ""}
    if "MANUAL_MONTH" not in config:
        config["MANUAL_MONTH"] = ""
    return config

def save_config():
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4, ensure_ascii=False)

config = load_config()

# 직종 리스트 (드롭다운 메뉴용)
DEPARTMENT_LIST = ["IT네트워크시스템", "클라우드컴퓨팅", "사이버보안", "공업전자기기", "메카트로닉스"]

# D-Day 날짜 설정 (기본)
DDAY_LOCAL = datetime(2025, 4, 7)  # '지방' 이벤트 날짜
DDAY_NATIONAL = datetime(2025, 9, 20)  # '전국' 이벤트 날짜

def adjusted_date(dt):
    return dt - timedelta(days=1) if dt.hour < 2 else dt

def gsn(dt):
    dt = adjusted_date(dt)
    # 수동 월 선택: 설정에 MANUAL_MONTH 값이 있으면 사용, 아니면 자동으로 현재 월 사용
    manual_month = config.get("MANUAL_MONTH", "").strip()
    if manual_month:
        month_str = manual_month
    else:
        month_str = f"{dt.month}월"
    manual_week = config.get("WEEK_NUMBER", "").strip()
    if not manual_week:
        tk.messagebox.showwarning("경고", "설정에서 째주를 선택하지 않았습니다.")
        manual_week = f"{(dt.day - 1) // 7 + 1}째주"
    return f"{month_str} {manual_week}"

def gti(dt):
    dt = adjusted_date(dt)
    return (dt.weekday() - 2) % 7

def ft(dt):
    period = "오전" if dt.hour < 12 else "오후"
    hour12 = dt.hour if 1 <= dt.hour <= 12 else (dt.hour - 12 if dt.hour > 12 else 12)
    return f"{dt.year}. {dt.month}. {dt.day} {period} {hour12:02d}:{dt.minute:02d}:{dt.second:02d}"

def calculate_dday(target_date):
    today = get_now().date()
    delta = target_date.date() - today
    days = delta.days
    if days == 0:
        return "D-DAY"
    elif days > 0:
        return f"D-{days}"
    else:
        return f"D+{-days}"

# 구글 스프레드시트 연결 (google.json 파일이 같은 폴더에 있어야 함)
try:
    gc = gspread.service_account("./google.json")
except FileNotFoundError:
    temp_root = tk.Tk()
    temp_root.withdraw()
    tk.messagebox.showerror("Error", "google.json 파일을 찾을 수 없습니다.\n실행 파일과 같은 위치에 넣어주세요!")
    sys.exit(1)

# 스프레드시트 URL 상수
SPREADSHEET_URL = "Google spreadsheet url"

def get_worksheet():
    """현재 설정된 시트 이름으로 시트를 불러옵니다. 없으면 None 반환."""
    now_actual = get_now()
    sheet_name = gsn(now_actual)
    try:
        ws = gc.open_by_url(SPREADSHEET_URL).worksheet(sheet_name)
        return ws
    except gspread.WorksheetNotFound:
        return None

def get_user_row(ws, checkin_col):
    """
    ws(워크시트)에서 '이름' 칼럼(checkin_col-2) 기준으로,
    config에 저장된 USER_NAME을 찾아 행 번호를 반환합니다.
    """
    name_col = checkin_col - 2
    user_name = config.get("USER_NAME", "").strip()
    if not user_name:
        tk.messagebox.showwarning("경고", "사용자 이름이 설정되어 있지 않습니다.\n설정을 확인해 주세요.")
        return None
    
    col_values = ws.col_values(name_col)
    for i, val in enumerate(col_values, start=1):
        if val.strip() == user_name:
            return i
    tk.messagebox.showwarning("경고", "시트에서 사용자의 이름을 찾을 수 없습니다.\n설정을 확인해 주세요.")
    return None

# 상수 설정
BASE_COL_CHECKIN = 3
BASE_COL_CHECKOUT = 4
BASE_COL_GOOUT = 5
TABLE_WIDTH = 6

def in_():
    now_actual = get_now()
    ws = get_worksheet()
    if ws is None:
        sheet_name = gsn(now_actual)
        tk.messagebox.showwarning("경고", f"시트 '{sheet_name}'가 존재하지 않습니다.\n설정을 확인하세요.")
        return
    formatted = ft(now_actual)
    table_idx = gti(now_actual)
    checkin_col = BASE_COL_CHECKIN + table_idx * TABLE_WIDTH
    row_number = get_user_row(ws, checkin_col)
    if row_number is None:
        return  # 이름이 없으므로 동작 중단
    cell_checkin = rowcol_to_a1(row_number, checkin_col)
    ws.update_acell(cell_checkin, formatted)
    dept_col = checkin_col - 1 
    cell_dept = rowcol_to_a1(row_number, dept_col)
    ws.update_acell(cell_dept, config["DEPARTMENT_NAME"])

def out():
    now_actual = get_now()
    ws = get_worksheet()
    if ws is None:
        sheet_name = gsn(now_actual)
        tk.messagebox.showwarning("경고", f"시트 '{sheet_name}'가 존재하지 않습니다.\n설정을 확인하세요.")
        return
    formatted = ft(now_actual)
    table_idx = gti(now_actual)
    checkin_col = BASE_COL_CHECKIN + table_idx * TABLE_WIDTH
    row_number = get_user_row(ws, checkin_col)
    if row_number is None:
        return  # 이름이 없으므로 동작 중단
    col = BASE_COL_CHECKOUT + table_idx * TABLE_WIDTH
    cell = rowcol_to_a1(row_number, col)
    ws.update_acell(cell, formatted)

def outside():
    out_start = get_now()
    ws = get_worksheet()
    if ws is None:
        sheet_name = gsn(out_start)
        tk.messagebox.showwarning("경고", f"시트 '{sheet_name}'가 존재하지 않습니다.\n설정을 확인하세요.")
        return
    start_str = out_start.strftime("%H:%M")
    out_window = tk.Toplevel(root)
    out_window.title("외출 측정")
    out_window.geometry("150x100")
    out_window.attributes('-topmost', True)
    label = tk.Label(out_window, text=f"외출 시작: {start_str}")
    label.pack(pady=10)
    
    def rfo():
        out_end = get_now()
        ws_inner = get_worksheet()
        if ws_inner is None:
            sheet_name_inner = gsn(out_end)
            tk.messagebox.showwarning("경고", f"시트 '{sheet_name_inner}'가 존재하지 않습니다.\n설정을 확인하세요.")
            out_window.destroy()
            return
        end_str = out_end.strftime("%H:%M")
        reason = tk.simpledialog.askstring("사유 입력", "외출 사유를 입력하세요:", parent=out_window)
        if reason is None:
            reason = ""
        result_str = f"{start_str}~{end_str}({reason})"
        table_idx_inner = gti(out_end)
        checkin_col_inner = BASE_COL_CHECKIN + table_idx_inner * TABLE_WIDTH
        row_number_inner = get_user_row(ws_inner, checkin_col_inner)
        if row_number_inner is None:
            out_window.destroy()
            return
        col = BASE_COL_GOOUT + table_idx_inner * TABLE_WIDTH
        cell = rowcol_to_a1(row_number_inner, col)
        # 기존 셀 내용 확인 후, 있다면 줄바꿈하여 추가
        existing_value = ws_inner.acell(cell).value
        if existing_value:
            new_value = existing_value + "\n" + result_str
        else:
            new_value = result_str
        ws_inner.update_acell(cell, new_value)
        out_window.destroy()
    
    btn_return = tk.Button(out_window, text="복귀", command=rfo)
    btn_return.pack(pady=10)

def open_settings():
    settings_window = tk.Toplevel(root)
    settings_window.title("설정")
    settings_window.geometry("300x450")
    settings_window.attributes('-topmost', True)
    
    tk.Label(settings_window, text="부서 선택:").pack(pady=5)
    department_var = tk.StringVar(settings_window)
    department_var.set(config["DEPARTMENT_NAME"])
    department_menu = tk.OptionMenu(settings_window, department_var, *DEPARTMENT_LIST)
    department_menu.pack(pady=5)
    
    tk.Label(settings_window, text="사용자 이름:").pack(pady=5)
    name_entry = tk.Entry(settings_window)
    name_entry.insert(0, config.get("USER_NAME", ""))
    name_entry.pack(pady=5)
    
    tk.Label(settings_window, text="째주 선택:").pack(pady=5)
    week_options = ["1째주", "2째주", "3째주", "4째주", "5째주", "6째주"]
    week_var = tk.StringVar(settings_window)
    if config.get("WEEK_NUMBER", "") in week_options:
        week_var.set(config["WEEK_NUMBER"])
    else:
        week_var.set(week_options[0])
    week_menu = tk.OptionMenu(settings_window, week_var, *week_options)
    week_menu.pack(pady=5)
    
    # 월 선택 옵션 추가 (자동 또는 수동 선택)
    tk.Label(settings_window, text="월 선택:").pack(pady=5)
    month_options = ["자동"] + [f"{i}월" for i in range(1, 13)]
    month_var = tk.StringVar(settings_window)
    current_manual_month = config.get("MANUAL_MONTH", "").strip()
    month_var.set(current_manual_month if current_manual_month else "자동")
    month_menu = tk.OptionMenu(settings_window, month_var, *month_options)
    month_menu.pack(pady=5)
    
    # 커스텀 디데이 설정 항목
    tk.Label(settings_window, text="커스텀 디데이 이벤트명:").pack(pady=5)
    custom_label_entry = tk.Entry(settings_window)
    custom_label_entry.insert(0, config.get("CUSTOM_DDAY", {}).get("label", ""))
    custom_label_entry.pack(pady=5)
    
    tk.Label(settings_window, text="커스텀 디데이 날짜 (YYYY-MM-DD):").pack(pady=5)
    custom_date_entry = tk.Entry(settings_window)
    custom_date_entry.insert(0, config.get("CUSTOM_DDAY", {}).get("date", ""))
    custom_date_entry.pack(pady=5)
    
    def save_settings():
        config["DEPARTMENT_NAME"] = department_var.get()
        config["USER_NAME"] = name_entry.get().strip()
        config["WEEK_NUMBER"] = week_var.get().strip()
        # 월 선택 저장: "자동"이면 빈 문자열로 저장
        selected_month = month_var.get()
        config["MANUAL_MONTH"] = "" if selected_month == "자동" else selected_month
        # 커스텀 디데이 저장
        config["CUSTOM_DDAY"]["label"] = custom_label_entry.get().strip()
        config["CUSTOM_DDAY"]["date"] = custom_date_entry.get().strip()
        save_config()
        update_dday_labels()  # 디데이 라벨 업데이트
        settings_window.destroy()
    
    save_button = tk.Button(settings_window, text="저장", command=save_settings)
    save_button.pack(pady=10)

def update_dday_labels():
    # 기본 이벤트 디데이
    dday_local_label.config(text=f"지방 {calculate_dday(DDAY_LOCAL)}")
    dday_national_label.config(text=f"전국 {calculate_dday(DDAY_NATIONAL)}")
    # 커스텀 디데이 이벤트 (설정되어 있으면)
    custom = config.get("CUSTOM_DDAY", {})
    label = custom.get("label", "").strip()
    date_str = custom.get("date", "").strip()
    if label and date_str:
        try:
            custom_date = datetime.strptime(date_str, "%Y-%m-%d")
            custom_text = f"{label} {calculate_dday(custom_date)}"
        except ValueError:
            custom_text = f"{label} (날짜 형식 오류)"
        custom_dday_label.config(text=custom_text)
    else:
        custom_dday_label.config(text="")

# 자동 리로드 기능 (매일 3시에 앱 재시작)
def schedule_reload():
    now = get_now()
    target = now.replace(hour=3, minute=0, second=0, microsecond=0)
    if now >= target:
        target += timedelta(days=1)
    delay_ms = int((target - now).total_seconds() * 1000)
    root.after(delay_ms, reload_app)

#Debug
#def schedule_reload():
#    root.after(5000, reload_app)

def reload_app():
    python_exe = sys.executable
    os.execl(python_exe, python_exe, *sys.argv)

# 메인 창 생성
root = tk.Tk()
root.title("근태 관리")
root.geometry("250x270")
root.attributes('-topmost', True)

bold_font = tkFont.Font(family="Helvetica", size=14, weight="bold")

btn_check_in = tk.Button(root, text="출근", command=in_, width=20)
btn_check_in.pack(padx=10, pady=5)

btn_check_out = tk.Button(root, text="퇴근", command=out, width=20)
btn_check_out.pack(padx=10, pady=5)

btn_go_out = tk.Button(root, text="외출", command=outside, width=20)
btn_go_out.pack(padx=10, pady=5)

btn_settings = tk.Button(root, text="설정", command=open_settings, width=20)
btn_settings.pack(padx=10, pady=5)

# 디데이 라벨들
dday_local_label = tk.Label(root, font=bold_font)
dday_local_label.pack(pady=5)

dday_national_label = tk.Label(root, font=bold_font)
dday_national_label.pack(pady=5)

custom_dday_label = tk.Label(root, font=bold_font, fg="blue")
custom_dday_label.pack(pady=5)

update_dday_labels()  # 시작시 라벨 업데이트
schedule_reload()      # 자동 리로드 예약

root.mainloop()
