import tkinter as tk
import tkinter.font as tkFont
from datetime import datetime, timedelta
import gspread
from gspread.utils import rowcol_to_a1

# 설정 변수
SERVICE_ACCOUNT_FILE = "./google.json"
SPREADSHEET_URL = "구글 시트 URL"
BASE_ROW = 25  # 사용자의 행 번호
DEPARTMENT_NAME = "IT네트워크시스템"  # 부서명
# D-Day 날짜 설정
DDAY_LOCAL = datetime(2025, 4, 7)  # '지방' 이벤트 날짜
DDAY_NATIONAL = datetime(2025, 9, 20)  # '전국' 이벤트 날짜

def gsn(dt):
    return f"{dt.month}월 {(dt.day - 1) // 7 + 1}째주"

def gti(dt):
    return (dt.weekday() - 2) % 7

def ft(dt):
    period = "오전" if dt.hour < 12 else "오후"
    hour12 = dt.hour if 1 <= dt.hour <= 12 else (dt.hour - 12 if dt.hour > 12 else 12)
    return f"{dt.year}. {dt.month}. {dt.day} {period} {hour12:02d}:{dt.minute:02d}:{dt.second:02d}"

gc = gspread.service_account(SERVICE_ACCOUNT_FILE)
now = datetime.now()
sheet_name = gsn(now)
try:
    worksheet = gc.open_by_url(SPREADSHEET_URL).worksheet(sheet_name)
except gspread.WorksheetNotFound:
    worksheet = gc.open_by_url(SPREADSHEET_URL).add_worksheet(title=sheet_name, rows="100", cols="50")

BASE_COL_CHECKIN = 3
BASE_COL_CHECKOUT = 4
BASE_COL_GOOUT = 5
TABLE_WIDTH = 6

def in_():
    now_actual = datetime.now()
    formatted = ft(now_actual)
    table_idx = gti(now_actual)
    col_checkin = BASE_COL_CHECKIN + table_idx * TABLE_WIDTH
    cell_checkin = rowcol_to_a1(BASE_ROW, col_checkin)
    worksheet.update_acell(cell_checkin, formatted)
    col_left = col_checkin - 1
    cell_left = rowcol_to_a1(BASE_ROW, col_left)
    worksheet.update_acell(cell_left, DEPARTMENT_NAME)

def out():
    now_actual = datetime.now()
    formatted = ft(now_actual)
    table_idx = gti(now_actual)
    col = BASE_COL_CHECKOUT + table_idx * TABLE_WIDTH
    cell = rowcol_to_a1(BASE_ROW, col)
    worksheet.update_acell(cell, formatted)

out_start = None

def outside():
    global out_start
    out_start = datetime.now()
    start_str = out_start.strftime("%H:%M")
    out_window = tk.Toplevel(root)
    out_window.title("외출 측정")
    out_window.geometry("300x150")
    out_window.attributes('-topmost', True)
    label = tk.Label(out_window, text=f"외출 시작: {start_str}")
    label.pack(pady=10)
    def rfo():
        out_end = datetime.now()
        end_str = out_end.strftime("%H:%M")
        reason = tk.simpledialog.askstring("사유 입력", "외출 사유를 입력하세요:", parent=out_window)
        if reason is None:
            reason = ""
        result_str = f"{start_str}~{end_str}({reason})"
        table_idx_inner = gti(out_end)
        col = BASE_COL_GOOUT + table_idx_inner * TABLE_WIDTH
        cell = rowcol_to_a1(BASE_ROW, col)
        worksheet.update_acell(cell, result_str)
        out_window.destroy()
    btn_return = tk.Button(out_window, text="복귀", command=rfo)
    btn_return.pack(pady=10)

# D-Day 계산 함수
def calculate_dday(target_date):
    today = datetime.now().date()
    delta = target_date.date() - today
    days = delta.days
    if days == 0:
        return "D-DAY"
    elif days > 0:
        return f"D-{days}"
    else:
        return f"D+{-days}"

# Tkinter 윈도우 생성
root = tk.Tk()
root.title("근태 관리")
root.geometry("250x200")
root.attributes('-topmost', True)

# 폰트 설정
bold_font = tkFont.Font(family="Helvetica", size=12, weight="bold")

# 출근 버튼
btn_check_in = tk.Button(root, text="출근", command=in_, width=20)
btn_check_in.pack(padx=10, pady=5)

# 퇴근 버튼
btn_check_out = tk.Button(root, text="퇴근", command=out, width=20)
btn_check_out.pack(padx=10, pady=5)

# 외출 버튼
btn_go_out = tk.Button(root, text="외출", command=outside, width=20)
btn_go_out.pack(padx=10, pady=5)

# D-Day 표시
dday_local_text = f"지방 {calculate_dday(DDAY_LOCAL)}"
dday_local_label = tk.Label(root, text=dday_local_text, font=bold_font)
dday_local_label.pack(pady=5)

dday_national_text = f"전국 {calculate_dday(DDAY_NATIONAL)}"
dday_national_label = tk.Label(root, text=dday_national_text, font=bold_font)
dday_national_label.pack(pady=5)

root.mainloop()
