import json
import os
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import openpyxl
from openpyxl.styles import Font
import urllib.request
import urllib.error
import subprocess
import sys
import threading

VERSION = "1.0.0"
GITHUB_USER = "xzn32"
GITHUB_REPO = "sim-inventory"
VERSION_URL  = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/main/version.json"

def check_for_updates(silent=False):
    def _check():
        try:
            status_label.config(text="Checking for updates...")
            root.update()

            req = urllib.request.Request(
                VERSION_URL,
                headers={"User-Agent": "SIM-Inventory-App", "Cache-Control": "no-cache"}
            )
            with urllib.request.urlopen(req, timeout=8) as resp:
                info = json.loads(resp.read().decode())

            latest  = info.get("version", VERSION)
            dl_url  = info.get("download_url", "")
            notes   = info.get("changelog", [])

            if latest > VERSION:
                notes_text = "\n".join(f"• {n}" for n in notes) if notes else "No details provided."
                answer = messagebox.askyesno(
                    "Update Available",
                    f"New version available: {latest}\n"
                    f"Your version:          {VERSION}\n\n"
                    f"What's new:\n{notes_text}\n\n"
                    "Download and install now?"
                )
                if answer:
                    download_and_install(dl_url, latest)
            else:
                if not silent:
                    messagebox.showinfo("Up to Date", f"You have the latest version ({VERSION}).")

        except urllib.error.URLError:
            if not silent:
                messagebox.showwarning(
                    "No Connection",
                    "Could not reach GitHub.\n"
                    "Make sure you are connected to your own WiFi, then try again."
                )
        except Exception as e:
            if not silent:
                messagebox.showerror("Update Error", f"Unexpected error:\n{str(e)}")
        finally:
            status_label.config(text="Ready")

    threading.Thread(target=_check, daemon=True).start()

def download_and_install(download_url, new_version):
    try:
        status_label.config(text=f"Downloading v{new_version}...")
        root.update()

        req = urllib.request.Request(
            download_url,
            headers={"User-Agent": "SIM-Inventory-App"}
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            new_code = resp.read().decode("utf-8")

        if not new_code.strip():
            messagebox.showerror("Update Failed", "Downloaded file is empty. Aborting.")
            return

        import shutil
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        script_backup = __file__ + f".backup_{ts}"
        shutil.copy2(__file__, script_backup)

        with open(__file__, "w", encoding="utf-8") as f:
            f.write(new_code)

        messagebox.showinfo(
            "Update Complete",
            f"Updated to v{new_version} successfully!\n"
            "The app will now restart."
        )
        restart_app()

    except Exception as e:
        messagebox.showerror("Update Failed", f"Could not install update:\n{str(e)}")
        status_label.config(text="Update failed")

def restart_app():
    save_data()
    subprocess.Popen([sys.executable, __file__])
    root.quit()
    root.destroy()
    sys.exit(0)

FILE = os.path.join(os.path.dirname(__file__), "sim_inventory.json")

def load_data():
    if not os.path.exists(FILE):
        return {"sims": [], "total_sold_ever": 0, "total_damaged_ever": 0}
    with open(FILE, "r") as f:
        loaded = json.load(f)
    if isinstance(loaded, list):
        for sim in loaded:
            sim.setdefault("type", "U Kurdistan")
            sim.setdefault("carrier", "Korek")
            sim.setdefault("status", "blank")
            sim.setdefault("phone_number", None)
            sim.setdefault("date_added", datetime.now().isoformat())
            sim.setdefault("date_sold", None)
        return {"sims": loaded, "total_sold_ever": 0, "total_damaged_ever": 0}
    else:
        loaded.setdefault("total_sold_ever", 0)
        loaded.setdefault("total_damaged_ever", 0)
        for sim in loaded["sims"]:
            sim.setdefault("type", "U Kurdistan")
            sim.setdefault("carrier", "Korek")
            sim.setdefault("status", "blank")
            sim.setdefault("phone_number", None)
            sim.setdefault("date_added", datetime.now().isoformat())
            sim.setdefault("date_sold", None)
        return loaded

def save_data():
    with open(FILE, "w") as f:
        json.dump(data, f, indent=4)

data  = load_data()
sims  = data["sims"]
existing = {sim["iccid"] for sim in sims}

def detect_type(iccid, sim_class):
    if iccid.startswith("892140800012"):
        region = "Kurdistan"
    elif iccid.startswith("892140800022"):
        region = "Baghdad"
    else:
        region = "Unknown"
    return f"{sim_class} {region}"

def update_counts():
    counts = {
        "U Kurdistan":0,"U Baghdad":0,"E Kurdistan":0,"E Baghdad":0,
        "U Unknown":0,"E Unknown":0,"Sold":0,"Damaged":0
    }
    for sim in sims:
        if sim["status"] == "sold":
            counts["Sold"] += 1
        elif sim["status"] == "damaged":
            counts["Damaged"] += 1
        else:
            counts.setdefault(sim["type"], 0)
            counts[sim["type"]] += 1
    for key, lbl in count_labels.items():
        if key in counts:
            lbl.config(text=f"{key}: {counts[key]}")
        elif key == "Total Sold Ever":
            lbl.config(text=f"Total Sold Ever: {data['total_sold_ever']}")
        elif key == "Total Damaged Ever":
            lbl.config(text=f"Total Damaged Ever: {data['total_damaged_ever']}")

def add_to_history(sim):
    history_tree.insert("", "end", values=(
        sim["id"], sim["iccid"], sim["type"], sim["carrier"], sim["status"],
        sim["phone_number"] or "",
        sim["date_added"].replace("T"," "),
        sim["date_sold"].replace("T"," ") if sim["date_sold"] else ""
    ))

def refresh_history():
    for row in history_tree.get_children():
        history_tree.delete(row)
    for sim in sims:
        add_to_history(sim)

def add_sim():
    iccid = iccid_entry.get().strip()
    sim_class = type_var.get()
    carrier = carrier_entry.get().strip() or "Korek"
    if not iccid or iccid in existing:
        iccid_entry.delete(0, tk.END)
        iccid_entry.focus_set()
        return
    sim_type = detect_type(iccid, sim_class)
    sim = {
        "id": len(sims)+1, "iccid": iccid, "type": sim_type,
        "carrier": carrier, "status": "blank", "phone_number": None,
        "date_added": datetime.now().isoformat(), "date_sold": None
    }
    sims.append(sim)
    existing.add(iccid)
    add_to_history(sim)
    iccid_entry.delete(0, tk.END)
    iccid_entry.focus_set()
    update_counts()
    save_data()

def bulk_add():
    bulk_text = simpledialog.askstring("Bulk Add SIMs", "Paste ICCIDs, one per line:")
    if not bulk_text:
        return
    iccids = [line.strip() for line in bulk_text.split('\n') if line.strip()]
    added = 0
    for iccid in iccids:
        if iccid and iccid not in existing:
            sim_type = detect_type(iccid, type_var.get())
            sim = {
                "id": len(sims)+1, "iccid": iccid, "type": sim_type,
                "carrier": carrier_entry.get().strip() or "Korek",
                "status": "blank", "phone_number": None,
                "date_added": datetime.now().isoformat(), "date_sold": None
            }
            sims.append(sim)
            existing.add(iccid)
            add_to_history(sim)
            added += 1
    if added:
        update_counts()
        save_data()
        messagebox.showinfo("Bulk Add", f"Added {added} SIMs.")
    else:
        messagebox.showinfo("Bulk Add", "No new SIMs added (duplicates or empty).")

def edit_sim():
    selected = history_tree.selection()
    if not selected:
        messagebox.showwarning("Edit SIM", "Please select a SIM to edit.")
        return
    iid = selected[0]
    sim_id = int(history_tree.item(iid, "values")[0])
    sim = next((s for s in sims if s["id"] == sim_id), None)
    if not sim:
        return

    edit_win = tk.Toplevel(root)
    edit_win.title("Edit SIM")
    edit_win.geometry("300x250")

    tk.Label(edit_win, text="ICCID:").grid(row=0, column=0, padx=5, pady=5)
    iccid_var = tk.StringVar(value=sim["iccid"])
    tk.Entry(edit_win, textvariable=iccid_var).grid(row=0, column=1, padx=5, pady=5)

    tk.Label(edit_win, text="Type:").grid(row=1, column=0, padx=5, pady=5)
    type_var_edit = tk.StringVar(value=sim["type"].split()[0] if " " in sim["type"] else sim["type"])
    tk.Entry(edit_win, textvariable=type_var_edit).grid(row=1, column=1, padx=5, pady=5)

    tk.Label(edit_win, text="Carrier:").grid(row=2, column=0, padx=5, pady=5)
    carrier_var_edit = tk.StringVar(value=sim["carrier"])
    tk.Entry(edit_win, textvariable=carrier_var_edit).grid(row=2, column=1, padx=5, pady=5)

    tk.Label(edit_win, text="Status:").grid(row=3, column=0, padx=5, pady=5)
    status_var_edit = tk.StringVar(value=sim["status"])
    tk.Entry(edit_win, textvariable=status_var_edit).grid(row=3, column=1, padx=5, pady=5)

    def save_edit():
        new_iccid   = iccid_var.get().strip()
        new_type    = type_var_edit.get().strip()
        new_carrier = carrier_var_edit.get().strip()
        new_status  = status_var_edit.get().strip()
        if new_iccid != sim["iccid"] and new_iccid in existing:
            messagebox.showerror("Error", "ICCID already exists.")
            return
        existing.discard(sim["iccid"])
        existing.add(new_iccid)
        sim["iccid"]   = new_iccid
        sim["type"]    = detect_type(new_iccid, new_type)
        sim["carrier"] = new_carrier
        sim["status"]  = new_status
        save_data()
        refresh_history()
        update_counts()
        edit_win.destroy()
        messagebox.showinfo("Edit SIM", "SIM updated successfully.")

    tk.Button(edit_win, text="Save", command=save_edit).grid(row=4, column=0, columnspan=2, pady=10)

def sell_sim():
    selected = history_tree.selection()
    if not selected:
        return
    for iid in selected:
        sim_id = int(history_tree.item(iid, "values")[0])
        for sim in sims:
            if sim["id"] == sim_id:
                sim["status"]    = "sold"
                sim["date_sold"] = datetime.now().isoformat()
                data["total_sold_ever"] += 1
                phone = simpledialog.askstring("Assign Number",
                    f"Assign phone number for ICCID {sim['iccid']}:")
                if phone:
                    sim["phone_number"] = phone
    save_data()
    refresh_history()
    update_counts()

def damage_sim():
    selected = history_tree.selection()
    if not selected:
        return
    for iid in selected:
        sim_id = int(history_tree.item(iid, "values")[0])
        for sim in sims:
            if sim["id"] == sim_id:
                sim["status"] = "damaged"
                data["total_damaged_ever"] += 1
    save_data()
    refresh_history()
    update_counts()

def delete_sim():
    selected = history_tree.selection()
    if not selected:
        return
    for iid in selected:
        sim_id = int(history_tree.item(iid, "values")[0])
        for i, sim in enumerate(sims):
            if sim["id"] == sim_id:
                sims.pop(i)
                existing.discard(sim["iccid"])
                break
    save_data()
    refresh_history()
    update_counts()

def reset_history():
    global sims, existing
    if messagebox.askyesno("Confirm Reset", "Are you sure you want to delete ALL SIM history?"):
        sims.clear()
        existing.clear()
        data["total_sold_ever"]   = 0
        data["total_damaged_ever"] = 0
        save_data()
        refresh_history()
        update_counts()

def search_sim():
    query = search_var.get().strip()
    for row in history_tree.get_children():
        history_tree.delete(row)
    for sim in sims:
        if query in sim["iccid"]:
            add_to_history(sim)

def export_unsold():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Unsold SIMs"
    headers = ["ID","ICCID","Type","Carrier","Status","Phone","Date Added","Date Sold"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for sim in sims:
        if sim["status"] != "sold":
            ws.append([
                sim["id"], sim["iccid"], sim["type"], sim["carrier"], sim["status"],
                sim["phone_number"] or "",
                sim["date_added"].replace("T"," "),
                sim["date_sold"].replace("T"," ") if sim["date_sold"] else ""
            ])
    save_path = os.path.join(os.path.dirname(__file__), "unsold_sims.xlsx")
    wb.save(save_path)
    messagebox.showinfo("Export Complete", f"Unsold SIMs exported to:\n{save_path}")

def import_from_excel():
    from tkinter import filedialog
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return
    try:
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        imported = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) < 2:
                continue
            iccid = str(row[1]).strip()
            if iccid and iccid not in existing:
                sim_type = detect_type(iccid, row[2] if len(row) > 2 and row[2] else "U")
                sim = {
                    "id": len(sims)+1, "iccid": iccid, "type": sim_type,
                    "carrier":      row[3] if len(row) > 3 and row[3] else "Korek",
                    "status":       row[4] if len(row) > 4 and row[4] else "blank",
                    "phone_number": row[5] if len(row) > 5 and row[5] else None,
                    "date_added":   row[6] if len(row) > 6 and row[6] else datetime.now().isoformat(),
                    "date_sold":    row[7] if len(row) > 7 and row[7] else None
                }
                sims.append(sim)
                existing.add(iccid)
                add_to_history(sim)
                imported += 1
        if imported:
            update_counts()
            save_data()
            messagebox.showinfo("Import Complete", f"Imported {imported} SIMs.")
        else:
            messagebox.showinfo("Import Complete", "No new SIMs imported (duplicates or invalid data).")
    except Exception as e:
        messagebox.showerror("Import Error", f"Error importing file:\n{str(e)}")

def backup_data():
    import shutil
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = os.path.join(os.path.dirname(__file__), f"sim_inventory_backup_{ts}.json")
    shutil.copy(FILE, backup_file)
    messagebox.showinfo("Backup Complete", f"Backup saved as:\n{backup_file}")

def show_summary():
    total   = len(sims)
    sold    = sum(1 for s in sims if s["status"] == "sold")
    damaged = sum(1 for s in sims if s["status"] == "damaged")
    blank   = sum(1 for s in sims if s["status"] == "blank")
    win = tk.Toplevel(root)
    win.title("Summary Statistics")
    win.geometry("250x180")
    for text in [
        f"Total SIMs: {total}",
        f"Sold: {sold}",
        f"Damaged: {damaged}",
        f"Blank / Available: {blank}",
        f"Total Sold Ever: {data['total_sold_ever']}",
        f"Total Damaged Ever: {data['total_damaged_ever']}"
    ]:
        tk.Label(win, text=text, font=("Arial", 11)).pack(pady=3)

def show_about():
    messagebox.showinfo("About",
        f"SIM Inventory Management System\n"
        f"Version: {VERSION}\n\n"
        f"Updates pulled from:\n"
        f"github.com/{GITHUB_USER}/{GITHUB_REPO}\n\n"
        f"© 2024 SIM Inventory"
    )

root = tk.Tk()
root.title(f"SIM Inventory v{VERSION}")
root.state("zoomed")
root.configure(bg="#ffffff")

style = ttk.Style()
style.theme_use("clam")
style.configure("TButton", font=("Arial", 10), padding=5, background="#f8f9fa",
                foreground="#212529", borderwidth=1, relief="flat")
style.map("TButton", background=[("active", "#e9ecef")])
style.configure("Primary.TButton", background="#007bff", foreground="white",
                font=("Arial", 10, "bold"), relief="raised", borderwidth=2)
style.map("Primary.TButton", background=[("active", "#0056b3")])
style.configure("Secondary.TButton", background="#6c757d", foreground="white", font=("Arial", 10))
style.map("Secondary.TButton", background=[("active", "#5a6268")])
style.configure("Update.TButton", background="#28a745", foreground="white",
                font=("Arial", 10, "bold"), relief="raised", borderwidth=2)
style.map("Update.TButton", background=[("active", "#1e7e34")])
style.configure("TLabel", font=("Arial", 10), background="#ffffff", foreground="#212529")
style.configure("TEntry", font=("Arial", 12), fieldbackground="#ffffff",
                foreground="#212529", insertcolor="#007bff")
style.configure("Treeview", font=("Arial", 10), rowheight=25,
                background="#ffffff", foreground="#212529", fieldbackground="#ffffff")
style.configure("Treeview.Heading", font=("Arial", 11, "bold"),
                background="#007bff", foreground="white")
style.map("Treeview", background=[("selected", "#007bff")])

frame_top = tk.Frame(root, bg="#007bff", padx=20, pady=10)
frame_top.pack(fill="x")
tk.Label(frame_top, text="SIM Inventory Management",
         font=("Arial", 18, "bold"), fg="white", bg="#007bff").pack(side="left")
ttk.Button(frame_top, text="⟳ Check for Updates",
           command=lambda: check_for_updates(silent=False),
           style="Update.TButton").pack(side="right", padx=5)
ttk.Button(frame_top, text="About",
           command=show_about,
           style="Secondary.TButton").pack(side="right", padx=5)

input_frame = tk.Frame(root, bg="#ffffff", padx=20, pady=10)
input_frame.pack(fill="x")

tk.Label(input_frame, text="ICCID:", font=("Arial", 12), bg="#ffffff").grid(row=0, column=0, padx=10, pady=5, sticky="e")
iccid_entry = tk.Entry(input_frame, font=("Arial", 12), bg="#ffffff", fg="#212529",
                        insertbackground="#007bff", width=25)
iccid_entry.grid(row=0, column=1, padx=10, pady=5)
iccid_entry.focus()
ttk.Button(input_frame, text="Add SIM", command=add_sim, style="Primary.TButton").grid(row=0, column=2, padx=5)
ttk.Button(input_frame, text="Bulk Add", command=bulk_add, style="Primary.TButton").grid(row=0, column=3, padx=5)

tk.Label(input_frame, text="Type:", font=("Arial", 12), bg="#ffffff").grid(row=0, column=4, padx=10, sticky="e")
type_var = tk.StringVar()
ttk.Combobox(input_frame, textvariable=type_var, values=["U","E"],
             state="readonly", font=("Arial", 12), width=5).grid(row=0, column=5, padx=5)
type_var.set("U")

tk.Label(input_frame, text="Carrier:", font=("Arial", 12), bg="#ffffff").grid(row=0, column=6, padx=10, sticky="e")
carrier_entry = tk.Entry(input_frame, font=("Arial", 12), bg="#ffffff", fg="#212529",
                          insertbackground="#007bff", width=15)
carrier_entry.grid(row=0, column=7, padx=5)
carrier_entry.insert(0, "Korek")

iccid_entry.bind("<Return>", lambda e: add_sim())

search_frame = tk.Frame(root, bg="#ffffff", padx=20, pady=5)
search_frame.pack(fill="x")
tk.Label(search_frame, text="Search ICCID:", font=("Arial", 12), bg="#ffffff").grid(row=0, column=0, padx=10, sticky="e")
search_var = tk.StringVar()
tk.Entry(search_frame, textvariable=search_var, font=("Arial", 12), bg="#ffffff",
         insertbackground="#007bff", width=25).grid(row=0, column=1, padx=10)
ttk.Button(search_frame, text="Search", command=search_sim,
           style="Secondary.TButton").grid(row=0, column=2, padx=10)

frame_counts = tk.Frame(root, bg="#ffffff", padx=20, pady=5)
frame_counts.pack(fill="x")
count_labels = {}
for i, key in enumerate(["U Kurdistan","U Baghdad","E Kurdistan","E Baghdad",
                          "U Unknown","E Unknown","Sold","Damaged",
                          "Total Sold Ever","Total Damaged Ever"]):
    lbl = tk.Label(frame_counts, text=f"{key}: 0", font=("Arial", 11), bg="#ffffff", fg="#212529")
    lbl.grid(row=0, column=i, padx=8)
    count_labels[key] = lbl

table_frame = tk.Frame(root, bg="#ffffff", padx=20, pady=10)
table_frame.pack(fill="both", expand=True)
columns = ("ID","ICCID","Type","Carrier","Status","Phone","Date Added","Date Sold")
history_tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=20)
for col in columns:
    history_tree.heading(col, text=col)
    history_tree.column(col, width=120, stretch=True)
history_tree.pack(fill="both", expand=True, side="left")
scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=history_tree.yview)
history_tree.configure(yscrollcommand=scrollbar.set)
scrollbar.pack(side="right", fill="y")

edit_entry  = None
edit_item   = None
edit_column = None

def start_inline_edit(event):
    global edit_entry, edit_item, edit_column
    region = history_tree.identify_region(event.x, event.y)
    if region != "cell":
        return
    column = history_tree.identify_column(event.x)
    item   = history_tree.identify_row(event.y)
    if not item or column == "#0":
        return
    col_index = int(column[1:]) - 1
    if col_index >= len(columns):
        return
    if columns[col_index] in ["ID","Date Added","Date Sold"]:
        return
    edit_item   = item
    edit_column = col_index
    current_value = history_tree.item(item, "values")[col_index]
    x, y, width, height = history_tree.bbox(item, column)
    edit_entry = tk.Entry(table_frame, font=("Arial", 10))
    edit_entry.insert(0, current_value)
    edit_entry.select_range(0, tk.END)
    edit_entry.focus()
    edit_entry.place(x=x, y=y, width=width, height=height)
    edit_entry.bind("<Return>",   save_inline_edit)
    edit_entry.bind("<Escape>",   cancel_inline_edit)
    edit_entry.bind("<FocusOut>", save_inline_edit)

def save_inline_edit(event=None):
    global edit_entry, edit_item, edit_column
    if not edit_entry or not edit_item:
        return
    new_value = edit_entry.get().strip()
    col_name  = columns[edit_column]
    sim_id    = int(history_tree.item(edit_item, "values")[0])
    sim       = next((s for s in sims if s["id"] == sim_id), None)
    if not sim:
        cancel_inline_edit()
        return
    if col_name == "ICCID":
        if new_value and new_value != sim["iccid"]:
            if new_value in existing:
                messagebox.showerror("Error", "ICCID already exists!")
                cancel_inline_edit()
                return
            existing.discard(sim["iccid"])
            existing.add(new_value)
            sim["iccid"] = new_value
    elif col_name == "Type":
        sim["type"] = detect_type(sim["iccid"], new_value)
    elif col_name == "Carrier":
        sim["carrier"] = new_value
    elif col_name == "Status":
        old_status = sim["status"]
        sim["status"] = new_value.lower()
        if old_status != "sold" and new_value.lower() == "sold":
            data["total_sold_ever"] += 1
        elif old_status != "damaged" and new_value.lower() == "damaged":
            data["total_damaged_ever"] += 1
    elif col_name == "Phone":
        sim["phone_number"] = new_value if new_value else None
    save_data()
    refresh_history()
    update_counts()
    cancel_inline_edit()

def cancel_inline_edit(event=None):
    global edit_entry
    if edit_entry:
        edit_entry.destroy()
        edit_entry = None

history_tree.bind("<Double-1>", start_inline_edit)

def copy_iccid():
    selected = history_tree.selection()
    if selected:
        root.clipboard_clear()
        root.clipboard_append(history_tree.item(selected[0], "values")[1])

def copy_row():
    selected = history_tree.selection()
    if selected:
        values = history_tree.item(selected[0], "values")
        root.clipboard_clear()
        root.clipboard_append("\t".join(str(v) for v in values))

def quick_sell():
    selected = history_tree.selection()
    if not selected:
        return
    for iid in selected:
        sim_id = int(history_tree.item(iid, "values")[0])
        for sim in sims:
            if sim["id"] == sim_id and sim["status"] != "sold":
                sim["status"]    = "sold"
                sim["date_sold"] = datetime.now().isoformat()
                data["total_sold_ever"] += 1
                phone = simpledialog.askstring("Assign Number",
                    f"Assign phone number for ICCID {sim['iccid']}:")
                if phone:
                    sim["phone_number"] = phone
    save_data(); refresh_history(); update_counts()

def quick_damage():
    selected = history_tree.selection()
    if not selected:
        return
    for iid in selected:
        sim_id = int(history_tree.item(iid, "values")[0])
        for sim in sims:
            if sim["id"] == sim_id and sim["status"] != "damaged":
                sim["status"] = "damaged"
                data["total_damaged_ever"] += 1
    save_data(); refresh_history(); update_counts()

def quick_assign_phone():
    selected = history_tree.selection()
    if not selected:
        messagebox.showwarning("Assign Phone", "Please select a SIM first.")
        return
    phone = simpledialog.askstring("Assign Phone Number", "Enter phone number:")
    if phone:
        for iid in selected:
            sim_id = int(history_tree.item(iid, "values")[0])
            for sim in sims:
                if sim["id"] == sim_id:
                    sim["phone_number"] = phone
        save_data(); refresh_history(); update_counts()

popup_menu = tk.Menu(root, tearoff=0)
popup_menu.add_command(label="Copy ICCID",           command=copy_iccid)
popup_menu.add_command(label="Copy Row",             command=copy_row)
popup_menu.add_separator()
popup_menu.add_command(label="Quick Edit",           command=edit_sim)
popup_menu.add_command(label="Mark as Sold",         command=quick_sell)
popup_menu.add_command(label="Mark as Damaged",      command=quick_damage)
popup_menu.add_command(label="Assign Phone Number",  command=quick_assign_phone)
popup_menu.add_separator()
popup_menu.add_command(label="Delete Selected",      command=delete_sim)

history_tree.bind("<Button-3>", lambda e: popup_menu.post(e.x_root, e.y_root))
history_tree.bind("<Control-c>", lambda e: copy_row())
root.bind("<Control-s>", lambda e: quick_sell())
root.bind("<Control-d>", lambda e: quick_damage())
root.bind("<Control-e>", lambda e: edit_sim())
root.bind("<Delete>",    lambda e: delete_sim())

button_frame = tk.Frame(root, bg="#ffffff", padx=20, pady=10)
button_frame.pack(fill="x")
for i, (text, cmd, btn_style) in enumerate([
    ("Edit Selected",        edit_sim,          "Primary.TButton"),
    ("Mark Sold",            sell_sim,          "Primary.TButton"),
    ("Mark Damaged",         damage_sim,        "Primary.TButton"),
    ("Delete Selected",      delete_sim,        "Primary.TButton"),
    ("Backup Data",          backup_data,       "Secondary.TButton"),
    ("Summary",              show_summary,      "Secondary.TButton"),
    ("Reset All",            reset_history,     "Secondary.TButton"),
    ("Import from Excel",    import_from_excel, "Secondary.TButton"),
    ("Export Unsold",        export_unsold,     "Secondary.TButton"),
]):
    ttk.Button(button_frame, text=text, command=cmd, style=btn_style).grid(row=0, column=i, padx=5, pady=5)

status_frame = tk.Frame(root, bg="#f8f9fa", padx=20, pady=5)
status_frame.pack(fill="x")
tk.Label(status_frame, text=f"Saving to: {FILE}",
         font=("Arial", 9), bg="#f8f9fa", fg="#6c757d").pack(side="left")
status_label = tk.Label(status_frame, text="Ready",
                         font=("Arial", 9), bg="#f8f9fa", fg="#6c757d")
status_label.pack(side="right")

def show_shortcuts():
    messagebox.showinfo("Keyboard Shortcuts & Tips", """
Keyboard Shortcuts:
• Double-click cell : Inline edit
• Right-click       : Context menu
• Ctrl+S            : Mark selected as Sold
• Ctrl+D            : Mark selected as Damaged
• Ctrl+E            : Edit selected SIM
• Delete            : Delete selected SIMs
• Ctrl+C            : Copy selected row

Inline Editing:
• Double-click any cell (except ID/Date fields)
• Press Enter or click outside to save
• Press Escape to cancel
""".strip())

ttk.Button(status_frame, text="?", width=3,
           command=show_shortcuts, style="Secondary.TButton").pack(side="right", padx=(0,10))

refresh_history()
update_counts()

root.after(3000, lambda: check_for_updates(silent=True))

root.mainloop()
