import tkinter as tk
from tkinter import simpledialog
from datetime import datetime, timedelta
import gspread
from gspread.utils import rowcol_to_a1

# 설정 변수
SERVICE_ACCOUNT_FILE = "./google.json"
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1p1lu3ytnzCgcLhUpDstoK8D0TmWPF5wLqIZWe_60Nq8/edit?usp=sharing"
BASE_ROW = 25  # 사용자 행 번호
DEPARTMENT_NAME = ""  # 직종명

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
        reason = simpledialog.askstring("사유 입력", "외출 사유를 입력하세요:", parent=out_window)
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

root = tk.Tk()
root.title("근태 관리")
root.geometry("300x200")
root.attributes('-topmost', True)
btn_check_in = tk.Button(root, text="출근", command=in_, width=20)
btn_check_in.pack(padx=10, pady=10)
btn_check_out = tk.Button(root, text="퇴근", command=out, width=20)
btn_check_out.pack(padx=10, pady=10)
btn_go_out = tk.Button(root, text="외출", command=outside, width=20)
btn_go_out.pack(padx=10, pady=10)
root.mainloop()
