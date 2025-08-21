import tkinter as tk
from tkinter import simpledialog, messagebox, ttk
from datetime import datetime, timedelta
from openpyxl import Workbook, load_workbook
import os
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

file_name = "study_log.xlsx"

# ---------------- Excel Setup ----------------
if not os.path.exists(file_name):
    wb = Workbook()

    # Logs
    ws1 = wb.active
    ws1.title = "Logs"
    ws1.append(["Date", "Start Time", "End Time", "Activity", "Duration"])

    # Summary
    ws2 = wb.create_sheet("Summary")
    ws2.append(["Date", "Total Study Time (HH:MM:SS)", "Change vs Previous Day"])

    # Daily Target
    ws_d = wb.create_sheet("DailyTarget")
    ws_d.append(["Date", "Target Hours", "Earned Hours", "Progress %"])

    # Weekly Target
    ws_w = wb.create_sheet("WeeklyTarget")
    ws_w.append(["Week Start", "Target Hours", "Earned Hours", "Progress %"])

    wb.save(file_name)

wb = load_workbook(file_name)
logs_ws = wb["Logs"]
summary_ws = wb["Summary"]
daily_ws = wb["DailyTarget"]
weekly_ws = wb["WeeklyTarget"]

start_time = None

# ---------------- Helpers ----------------
def to_td(timestr):
    if not timestr:
        return timedelta()
    h, m, s = str(timestr).split(":")
    return timedelta(hours=int(h), minutes=int(m), seconds=int(float(s)))

def get_week_start(date_obj):
    return (date_obj - timedelta(days=date_obj.weekday())).date()

def add_or_update_target(ws, key, target_hours):
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=False), start=2):
        if row[0].value == key:
            ws.cell(i, 2).value = target_hours
            return
    row_num = ws.max_row + 1
    ws.cell(row_num, 1).value = key
    ws.cell(row_num, 2).value = target_hours
    ws.cell(row_num, 3).value = 0
    ws.cell(row_num, 4).value = "0%"

def add_time(ws, key, duration_hours):
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=False), start=2):
        if row[0].value == key:
            prev = row[2].value or 0
            target = row[1].value or 0
            new_val = prev + duration_hours
            percent = min(100, round((new_val / target) * 100, 1)) if target else 0
            ws.cell(i, 3).value = round(new_val, 2)
            ws.cell(i, 4).value = f"{percent}%"
            return
    row_num = ws.max_row + 1
    ws.cell(row_num, 1).value = key
    ws.cell(row_num, 2).value = 0
    ws.cell(row_num, 3).value = duration_hours
    ws.cell(row_num, 4).value = "0%"

# ---------------- Session ----------------
def start_session():
    global start_time
    start_time = datetime.now()
    status_label.config(text=f"Started at {start_time.strftime('%H:%M:%S')}")

def stop_session():
    global start_time
    if not start_time:
        messagebox.showwarning("Error", "Start the session first!")
        return

    end_time = datetime.now()
    duration = end_time - start_time
    today_str = str(datetime.today().date())

    activity = simpledialog.askstring("Activity", "What did you study? (Optional)") or ""

    logs_ws.append([today_str,
                    start_time.strftime("%H:%M:%S"),
                    end_time.strftime("%H:%M:%S"),
                    activity,
                    str(duration).split(".")[0]])

    # Summary
    found = False
    for i, row in enumerate(summary_ws.iter_rows(min_row=2, values_only=False), start=2):
        if row[0].value == today_str:
            old_td = to_td(row[1].value)
            new_total = old_td + duration
            summary_ws.cell(i, 2).value = str(new_total).split(".")[0]
            found = True
            break
    if not found:
        summary_ws.append([today_str, str(duration).split(".")[0], ""])

    # Targets
    hours = duration.total_seconds() / 3600
    add_time(daily_ws, today_str, hours)
    add_time(weekly_ws, str(get_week_start(datetime.today())), hours)

    wb.save(file_name)
    status_label.config(text=f"Saved! Duration: {str(duration).split('.')[0]}")
    start_time = None

# ---------------- Dashboard ----------------
def open_dashboard():
    dash = tk.Toplevel(root)
    dash.title("Dashboard")
    dash.geometry("600x500")

    notebook = ttk.Notebook(dash)
    notebook.pack(expand=True, fill="both")

    # Summary Tab
    frame_summary = tk.Frame(notebook)
    notebook.add(frame_summary, text="Summary")
    tree = ttk.Treeview(frame_summary, columns=("Date","Total","Change"), show="headings")
    for c in ("Date","Total","Change"):
        tree.heading(c, text=c)
    tree.pack(expand=True, fill="both")
    for row in summary_ws.iter_rows(min_row=2, values_only=True):
        tree.insert("", "end", values=row)

    # Daily Target Tab
    frame_daily = tk.Frame(notebook)
    notebook.add(frame_daily, text="Daily Target")
    tk.Button(frame_daily, text="Set Daily Target", command=set_daily_target).pack(pady=5)
    refresh_target_ui(frame_daily, daily_ws, str(datetime.today().date()))

    # Weekly Target Tab
    frame_weekly = tk.Frame(notebook)
    notebook.add(frame_weekly, text="Weekly Target")
    tk.Button(frame_weekly, text="Set Weekly Target", command=set_weekly_target).pack(pady=5)
    refresh_target_ui(frame_weekly, weekly_ws, str(get_week_start(datetime.today())))

    # Visualization Tab
    frame_vis = tk.Frame(notebook)
    notebook.add(frame_vis, text="Visualization")
    plot_summary(frame_vis)

def set_daily_target():
    t = simpledialog.askinteger("Daily Target", "Enter target hours today:")
    if t:
        add_or_update_target(daily_ws, str(datetime.today().date()), t)
        wb.save(file_name)

def set_weekly_target():
    t = simpledialog.askinteger("Weekly Target", "Enter target hours this week:")
    if t:
        add_or_update_target(weekly_ws, str(get_week_start(datetime.today())), t)
        wb.save(file_name)

def refresh_target_ui(parent, ws, key):
    for row in ws.iter_rows(min_row=2, values_only=True):
        if str(row[0]) == str(key):
            tk.Label(parent, text=f"Target: {row[1]} hrs").pack()
            tk.Label(parent, text=f"Earned: {row[2]} hrs").pack()
            pct = float(str(row[3]).strip('%')) if row[3] else 0
            bar = ttk.Progressbar(parent, length=400, mode="determinate", maximum=100)
            bar["value"] = pct
            bar.pack(pady=5)
            tk.Label(parent, text=f"{pct}%").pack()
            return
    tk.Label(parent, text="No target set yet.").pack()

def plot_summary(parent):
    dates, totals = [], []
    for row in summary_ws.iter_rows(min_row=2, values_only=True):
        if row[0] and row[1]:
            dates.append(row[0])
            h,m,s = map(int,str(row[1]).split(":"))
            totals.append(h + m/60 + s/3600)

    if not dates: 
        tk.Label(parent, text="No data yet").pack()
        return

    fig = Figure(figsize=(5,3))
    ax = fig.add_subplot(111)
    ax.plot(dates, totals, marker="o")
    ax.set_title("Daily Study Hours")
    ax.set_ylabel("Hours")
    ax.set_xlabel("Date")
    ax.tick_params(axis='x', rotation=45)

    canvas = FigureCanvasTkAgg(fig, master=parent)
    canvas.draw()
    canvas.get_tk_widget().pack(expand=True, fill="both")

# ---------------- GUI ----------------
root = tk.Tk()
root.title("Study Tracker")
root.geometry("250x200")

start_btn = tk.Button(root, text="Start", command=start_session, bg="green", fg="white", width=15, height=2)
start_btn.pack(pady=10)

stop_btn = tk.Button(root, text="Stop", command=stop_session, bg="red", fg="white", width=15, height=2)
stop_btn.pack(pady=10)

dash_btn = tk.Button(root, text="Dashboard", command=open_dashboard, bg="blue", fg="white", width=15, height=2)
dash_btn.pack(pady=10)

status_label = tk.Label(root, text="Click Start to begin", fg="blue")
status_label.pack(pady=10)

root.mainloop()
