import tkinter as tk
from tkinter import simpledialog, messagebox, ttk
from datetime import datetime, timedelta
from openpyxl import Workbook, load_workbook
import os

file_name = "study_log.xlsx"

# Create Excel file with two sheets if not exists
if not os.path.exists(file_name):
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Logs"
    ws1.append(["Date", "Start Time", "End Time", "Activity", "Duration"])

    ws2 = wb.create_sheet("Summary")
    ws2.append(["Date", "Total Study Time (HH:MM:SS)", "Change vs Previous Day"])
    wb.save(file_name)

wb = load_workbook(file_name)
logs_ws = wb["Logs"]
summary_ws = wb["Summary"]

start_time = None

# --- Helper: convert time string to timedelta ---
def to_td(timestr):
    if not timestr:
        return timedelta()
    parts = str(timestr).split(":")
    if len(parts) == 3:
        h, m, s = parts
        return timedelta(hours=int(h), minutes=int(m), seconds=int(float(s)))
    return timedelta()


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
    today = str(datetime.today().date())

    # Ask optional activity
    activity = simpledialog.askstring("Activity", "What did you study? (Optional)")
    if activity is None:
        activity = ""

    # Save to Logs sheet
    logs_ws.append([
        today,
        start_time.strftime("%H:%M:%S"),
        end_time.strftime("%H:%M:%S"),
        activity,
        str(duration).split(".")[0]
    ])

    # --- Update Summary ---
    found = False
    for i, row in enumerate(summary_ws.iter_rows(min_row=2, values_only=False), start=2):
        if row[0].value == today:
            old_total = row[1].value or "0:00:00"
            old_td = to_td(old_total)
            new_total = old_td + duration
            summary_ws.cell(i, 1).value = today
            summary_ws.cell(i, 2).value = str(new_total).split(".")[0]
            found = True

            # update comparison with previous day
            if i > 2:
                prev_td = to_td(summary_ws.cell(i-1, 2).value)
                diff = new_total - prev_td
                summary_ws.cell(i, 3).value = f"{'+' if diff.total_seconds()>=0 else '-'}{str(abs(diff)).split('.')[0]}"
            break

    if not found:
        new_total = duration
        row_num = summary_ws.max_row + 1
        summary_ws.cell(row_num, 1).value = today
        summary_ws.cell(row_num, 2).value = str(new_total).split(".")[0]

        # comparison with previous day if exists
        if row_num > 2:
            prev_td = to_td(summary_ws.cell(row_num-1, 2).value)
            diff = new_total - prev_td
            summary_ws.cell(row_num, 3).value = f"{'+' if diff.total_seconds()>=0 else '-'}{str(abs(diff)).split('.')[0]}"

    wb.save(file_name)

    status_label.config(text=f"Session saved! Duration: {str(duration).split('.')[0]}")
    start_time = None

    # auto refresh summary window if open
    if summary_win and summary_win.winfo_exists():
        refresh_summary()


def refresh_summary():
    """Refresh table contents"""
    for row in tree.get_children():
        tree.delete(row)

    for row in summary_ws.iter_rows(min_row=2, values_only=True):
        tree.insert("", "end", values=row)


def view_summary():
    global summary_win, tree
    summary_win = tk.Toplevel(root)
    summary_win.title("Study Summary")
    summary_win.geometry("450x250")

    tree = ttk.Treeview(summary_win, columns=("Date", "Total", "Change"), show="headings")
    tree.heading("Date", text="Date")
    tree.heading("Total", text="Total Study Time")
    tree.heading("Change", text="Vs Previous Day")

    tree.pack(expand=True, fill="both")
    refresh_summary()


# ---------------- GUI ----------------
root = tk.Tk()
root.title("Study Tracker")
root.geometry("350x250")

start_btn = tk.Button(root, text="Start", command=start_session,
                      bg="green", fg="white", width=15, height=2)
start_btn.pack(pady=10)

stop_btn = tk.Button(root, text="Stop", command=stop_session,
                     bg="red", fg="white", width=15, height=2)
stop_btn.pack(pady=10)

summary_btn = tk.Button(root, text="View Summary", command=view_summary,
                        bg="blue", fg="white", width=15, height=2)
summary_btn.pack(pady=10)

status_label = tk.Label(root, text="Click Start to begin", fg="blue")
status_label.pack(pady=10)

summary_win = None
tree = None

root.mainloop()
