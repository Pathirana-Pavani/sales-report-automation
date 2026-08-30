import os
import shutil
import time
import queue
import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from watchdog.observers.polling import PollingObserver as Observer
from watchdog.events import FileSystemEventHandler

from report_builder import DISTRICTS, build_excel_report

DEBOUNCE_SECONDS = 3

class DownloadHandler(FileSystemEventHandler):
    def __init__(self, file_queue):
        self.file_queue = file_queue
        self.recent_files = {}  # file_path -> last-seen timestamp
    def trigger_queue(self, file_path):
        if not file_path.lower().endswith('.csv'): return
        now = time.time()
        last_seen = self.recent_files.get(file_path)
        if last_seen is not None and (now - last_seen) < DEBOUNCE_SECONDS:
            return
        self.recent_files[file_path] = now
        time.sleep(0.5)
        self.file_queue.put(file_path)
    def on_created(self, event):
        if not event.is_directory: self.trigger_queue(event.src_path)
    def on_moved(self, event):
        if not event.is_directory: self.trigger_queue(event.dest_path)

class AppUI:
    def __init__(self, root):
        self.root = root
        self.root.withdraw() 
        self.file_queue = queue.Queue()
        self.target_base_dir = ""
        self.selected_region = ""
        self.setup_environment()
        self.build_control_panel()
        self.downloads_folder = os.path.expanduser("~/Downloads")
        self.observer = Observer()
        self.observer.schedule(DownloadHandler(self.file_queue), self.downloads_folder, recursive=False)
        self.observer.start()
        self.check_queue()

    def setup_environment(self):
        self.target_base_dir = filedialog.askdirectory(title="Select the Target Base Folder")
        if not self.target_base_dir: os._exit(0)
            
        setup_win = tk.Toplevel(self.root)
        setup_win.title("Select Region")
        setup_win.geometry("300x150")
        setup_win.attributes('-topmost', True)
        tk.Label(setup_win, text="Select the Region for this session:").pack(pady=10)
        region_var = tk.StringVar()
        region_cb = ttk.Combobox(setup_win, textvariable=region_var, values=list(DISTRICTS.keys()), state="readonly")
        region_cb.pack(pady=5)
        region_cb.current(0)
        
        def confirm_setup():
            self.selected_region = region_var.get()
            self.create_folder_structure()
            setup_win.destroy()
            
        tk.Button(setup_win, text="Confirm", command=confirm_setup).pack(pady=10)
        self.root.wait_window(setup_win)

    def create_folder_structure(self):
        base = os.path.join(self.target_base_dir, "Sales Report", self.selected_region)
        folders = ["2G Voice", "3G Voice", "3G Data", "4G Voice", "4G Data"]
        for f in folders: os.makedirs(os.path.join(base, f), exist_ok=True)

    def build_control_panel(self):
        self.root.deiconify() 
        self.root.title("Hutch KPI Monitor")
        self.root.geometry("400x250")
        self.root.attributes('-topmost', True)
        banner_text = "Do not rename your downloads!\nJust download files one by one - I will\nrename, move, and convert them automatically."
        tk.Label(self.root, text=banner_text, font=("Courier", 10, "bold"), bg="#005837", fg="white", pady=10, padx=10).pack(fill=tk.X)
        tk.Label(self.root, text=f"Monitoring Downloads for: {self.selected_region}", font=("Arial", 12, "bold")).pack(pady=20)
        tk.Label(self.root, text="Leave this open while you download CSVs.\nPopups will appear automatically.", fg="gray").pack()
        
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=20)
        
        tk.Button(btn_frame, text="Generate Final Report & Exit", command=self.generate_final_report, bg="blue", fg="white").pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="Exit Without Report", command=self.close_app, bg="red", fg="white").pack(side=tk.LEFT, padx=10)

    def check_queue(self):
        try:
            filepath = self.file_queue.get_nowait()
            self.prompt_file_classification(filepath)
        except queue.Empty: pass
        finally: self.root.after(1000, self.check_queue)

    def prompt_file_classification(self, filepath):
        filename = os.path.basename(filepath)
        popup = tk.Toplevel(self.root)
        popup.title("New KPI Data Detected!")
        popup.geometry("500x520")
        popup.attributes('-topmost', True)
        
        banner_text = "EASY FOCUS: Do not rename your downloads!\nSelect the details below by clicking the buttons."
        tk.Label(popup, text=banner_text, font=("Arial", 10, "bold"), bg="#0056b3", fg="white", pady=10, padx=10).pack(fill=tk.X)
        tk.Label(popup, text=f"Downloaded: {filename}", wraplength=450, fg="blue", pady=10).pack()
        
        tech_var, type_var, dist_var = tk.StringVar(), tk.StringVar(), tk.StringVar()

        def update_styles(var, btns_dict):
            selected = var.get()
            for opt, btn in btns_dict.items():
                if btn.cget('state') == tk.DISABLED: continue
                if opt == selected: btn.config(bg="#28a745", fg="white", relief="sunken")
                else: btn.config(bg="#e0e0e0", fg="black", relief="raised")

        def on_tech_click(val):
            tech_var.set(val)
            update_styles(tech_var, tech_btns)
            if val == "2G":
                type_var.set("Voice")
                type_btns["Data"].config(state=tk.DISABLED, bg="#d3d3d3", fg="gray", relief="flat")
                update_styles(type_var, type_btns)
            else:
                type_btns["Data"].config(state=tk.NORMAL)
                update_styles(type_var, type_btns)

        def on_type_click(val):
            type_var.set(val)
            update_styles(type_var, type_btns)

        def on_dist_click(val):
            dist_var.set(val)
            update_styles(dist_var, dist_btns)

        tk.Label(popup, text="Select Technology:", font=("Arial", 10, "bold")).pack(pady=(5, 0))
        tech_frame = tk.Frame(popup)
        tech_frame.pack()
        tech_btns = {}
        for opt in ["2G", "3G", "4G"]:
            btn = tk.Button(tech_frame, text=opt, width=12, font=("Arial", 10, "bold"), command=lambda o=opt: on_tech_click(o))
            btn.pack(side=tk.LEFT, padx=5, pady=5); tech_btns[opt] = btn

        tk.Label(popup, text="Select Traffic Type:", font=("Arial", 10, "bold")).pack(pady=(10, 0))
        type_frame = tk.Frame(popup)
        type_frame.pack()
        type_btns = {}
        for opt in ["Voice", "Data"]:
            btn = tk.Button(type_frame, text=opt, width=12, font=("Arial", 10, "bold"), command=lambda o=opt: on_type_click(o))
            btn.pack(side=tk.LEFT, padx=5, pady=5); type_btns[opt] = btn

        tk.Label(popup, text="Select Location/District:", font=("Arial", 10, "bold")).pack(pady=(10, 0))
        dist_frame = tk.Frame(popup)
        dist_frame.pack()
        dist_btns = {}
        row, col = 0, 0
        for opt in DISTRICTS[self.selected_region]:
            btn = tk.Button(dist_frame, text=opt, width=12, font=("Arial", 10, "bold"), command=lambda o=opt: on_dist_click(o))
            btn.grid(row=row, column=col, padx=5, pady=5); dist_btns[opt] = btn
            col += 1
            if col > 2: col = 0; row += 1

        update_styles(tech_var, tech_btns); update_styles(type_var, type_btns); update_styles(dist_var, dist_btns)

        def process_file():
            if not tech_var.get() or not type_var.get() or not dist_var.get():
                messagebox.showerror("Error", "Please select all fields.", parent=popup); return
            new_name_base = f"{tech_var.get()}_{type_var.get()}_{dist_var.get()}"
            folder_name = f"{tech_var.get()} {type_var.get()}"
            target_folder = os.path.join(self.target_base_dir, "Sales Report", self.selected_region, folder_name)
            csv_dest = os.path.join(target_folder, f"{new_name_base}.csv")
            excel_dest = os.path.join(target_folder, f"{new_name_base}.xlsx")
            try:
                shutil.move(filepath, csv_dest)
                df = pd.read_csv(csv_dest)
                df.to_excel(excel_dest, index=False); popup.destroy()
            except Exception as e:
                messagebox.showerror("File Error", f"Failed: {str(e)}", parent=popup)

        tk.Button(popup, text="Save & Convert", font=("Arial", 11, "bold"), command=process_file, bg="green", fg="white", width=20, pady=5).pack(pady=20)

    def generate_final_report(self):
        self.observer.stop(); self.root.withdraw()
        try:
            build_excel_report(self.target_base_dir, self.selected_region)
        except PermissionError:
            messagebox.showerror("File is Open", "Close the report file!"); self.root.deiconify(); self.observer.start(); return
        messagebox.showinfo("Success", "Report Generated!"); self.close_app()

    def close_app(self):
        try: self.observer.stop()
        except: pass
        self.root.destroy(); os._exit(0)

if __name__ == "__main__":
    root = tk.Tk()
    app = AppUI(root)
    try: root.mainloop()
    except KeyboardInterrupt: app.close_app()