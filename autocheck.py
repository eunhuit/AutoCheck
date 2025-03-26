import tkinter as tk
import tkinter.font as tkFont
import tkinter.simpledialog
import tkinter.messagebox 
import json
import sys
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
            return json.load(f)
    except FileNotFoundError:
        return {
            "BASE_ROW": 25,
            "DEPARTMENT_NAME": "IT네트워크시스템",
            "USER_NAME": ""  # 기본 사용자 이름은 빈 문자열
        }

def save_config():
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4, ensure_ascii=False)

config = load_config()

# 직종 리스트 (드롭다운 메뉴용)
DEPARTMENT_LIST = ["IT네트워크시스템", "클라우드컴퓨팅", "사이버보안", "공업전자기기", "메카트로닉스"]

# D-Day 날짜 설정
DDAY_LOCAL = datetime(2025, 4, 7)  # '지방' 이벤트 날짜
DDAY_NATIONAL = datetime(2025, 9, 20)  # '전국' 이벤트 날짜

def adjusted_date(dt):
    return dt - timedelta(days=1) if dt.hour < 2 else dt

def gsn(dt):
    dt = adjusted_date(dt)
    return f"{dt.month}월 {(dt.day - 1) // 7 + 1}째주"

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

try:
    gc = gspread.service_account("./google.json")
except FileNotFoundError:
    temp_root = tk.Tk()
    temp_root.withdraw()
    tk.messagebox.showerror("Error", "google.json 파일을 찾을 수 없습니다.\n실행 파일과 같은 위치에 넣어주세요!")
    sys.exit(1)

now = get_now()
sheet_name = gsn(now)
try:
    worksheet = gc.open_by_url("구글 시트 URL").worksheet(sheet_name)
except gspread.WorksheetNotFound:
    worksheet = gc.open_by_url("구글 시트 URL").add_worksheet(title=sheet_name, rows="100", cols="50")

# 상수 설정
BASE_COL_CHECKIN = 3
BASE_COL_CHECKOUT = 4
BASE_COL_GOOUT = 5
TABLE_WIDTH = 6

def get_user_row(checkin_col):
    """
    출근 날짜가 입력되는 칼럼(checkin_col) 기준 왼쪽 2칸에 위치한 '이름' 칼럼에서,
    config에 저장된 USER_NAME을 찾아 행 번호를 반환합니다.
    이름을 찾지 못하면 경고창을 띄우고 None을 반환합니다.
    """
    name_col = checkin_col - 2
    user_name = config.get("USER_NAME", "").strip()
    if not user_name:
        tk.messagebox.showwarning("경고", "사용자 이름이 설정되어 있지 않습니다.\n설정을 확인해 주세요.")
        return None
    
    col_values = worksheet.col_values(name_col)
    for i, val in enumerate(col_values, start=1):
        if val.strip() == user_name:
            return i
    tk.messagebox.showwarning("경고", "시트에서 사용자의 이름을 찾을 수 없습니다.\n설정을 확인해 주세요.")
    return None

def in_():
    now_actual = get_now()
    formatted = ft(now_actual)
    table_idx = gti(now_actual)
    checkin_col = BASE_COL_CHECKIN + table_idx * TABLE_WIDTH
    row_number = get_user_row(checkin_col)
    if row_number is None:
        return  # 이름이 없으므로 동작 중단
    cell_checkin = rowcol_to_a1(row_number, checkin_col)
    worksheet.update_acell(cell_checkin, formatted)
    dept_col = checkin_col - 1 
    cell_dept = rowcol_to_a1(row_number, dept_col)
    worksheet.update_acell(cell_dept, config["DEPARTMENT_NAME"])

def out():
    now_actual = get_now()
    formatted = ft(now_actual)
    table_idx = gti(now_actual)
    checkin_col = BASE_COL_CHECKIN + table_idx * TABLE_WIDTH
    row_number = get_user_row(checkin_col)
    if row_number is None:
        return  # 이름이 없으므로 동작 중단
    col = BASE_COL_CHECKOUT + table_idx * TABLE_WIDTH
    cell = rowcol_to_a1(row_number, col)
    worksheet.update_acell(cell, formatted)

def outside():
    out_start = get_now()
    start_str = out_start.strftime("%H:%M")
    out_window = tk.Toplevel(root)
    out_window.title("외출 측정")
    out_window.geometry("150x100")
    out_window.attributes('-topmost', True)
    label = tk.Label(out_window, text=f"외출 시작: {start_str}")
    label.pack(pady=10)
    def rfo():
        out_end = get_now()
        end_str = out_end.strftime("%H:%M")
        reason = tk.simpledialog.askstring("사유 입력", "외출 사유를 입력하세요:", parent=out_window)
        if reason is None:
            reason = ""
        result_str = f"{start_str}~{end_str}({reason})"
        table_idx_inner = gti(out_end)
        checkin_col_inner = BASE_COL_CHECKIN + table_idx_inner * TABLE_WIDTH
        row_number_inner = get_user_row(checkin_col_inner)
        if row_number_inner is None:
            out_window.destroy()
            return  # 이름이 없으므로 동작 중단
        col = BASE_COL_GOOUT + table_idx_inner * TABLE_WIDTH
        cell = rowcol_to_a1(row_number_inner, col)
        worksheet.update_acell(cell, result_str)
        out_window.destroy()
    btn_return = tk.Button(out_window, text="복귀", command=rfo)
    btn_return.pack(pady=10)

def open_settings():
    settings_window = tk.Toplevel(root)
    settings_window.title("설정")
    settings_window.geometry("300x250")
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
    
    def save_settings():
        config["DEPARTMENT_NAME"] = department_var.get()
        config["USER_NAME"] = name_entry.get().strip()
        save_config()
        settings_window.destroy()
    
    save_button = tk.Button(settings_window, text="저장", command=save_settings)
    save_button.pack(pady=10)

root = tk.Tk()
root.title("근태 관리")
root.geometry("250x250")
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

dday_local_label = tk.Label(root, text=f"지방 {calculate_dday(DDAY_LOCAL)}", font=bold_font)
dday_local_label.pack(pady=5)

dday_national_label = tk.Label(root, text=f"전국 {calculate_dday(DDAY_NATIONAL)}", font=bold_font)
dday_national_label.pack(pady=5)

root.mainloop()
