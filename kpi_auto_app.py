import queue
import threading
from datetime import datetime, timedelta

import requests
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

import report_full_1 as core
import kpi_fetch


class KpiAutoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Hutch KPI Auto Report")
        self.root.geometry("560x560")
        self.log_queue = queue.Queue()
        self.target_dir = ""
        self._build_ui()
        self._poll_log()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 4}

        tk.Label(self.root, text="Hutch KPI Auto Report", font=("Arial", 14, "bold")).pack(pady=(10, 0))
        tk.Label(self.root, text="Fetches KPI data directly from the portal and builds the Excel report.",
                 fg="gray").pack()

        form = tk.Frame(self.root)
        form.pack(fill=tk.X, **pad)

        tk.Label(form, text="Username / Email:").grid(row=0, column=0, sticky="w")
        self.email_var = tk.StringVar()
        tk.Entry(form, textvariable=self.email_var, width=35).grid(row=0, column=1, sticky="w")

        tk.Label(form, text="Password:").grid(row=1, column=0, sticky="w")
        self.password_var = tk.StringVar()
        tk.Entry(form, textvariable=self.password_var, width=35, show="*").grid(row=1, column=1, sticky="w")

        tk.Label(form, text="Region:").grid(row=2, column=0, sticky="w")
        self.region_var = tk.StringVar(value="South")
        ttk.Combobox(form, textvariable=self.region_var, values=list(kpi_fetch.DISTRICTS.keys()),
                     state="readonly", width=32).grid(row=2, column=1, sticky="w")

        today = datetime.now()
        three_months_ago = today - timedelta(days=90)

        tk.Label(form, text="Start Date (YYYY-MM-DD):").grid(row=3, column=0, sticky="w")
        self.start_var = tk.StringVar(value=three_months_ago.strftime("%Y-%m-%d"))
        tk.Entry(form, textvariable=self.start_var, width=35).grid(row=3, column=1, sticky="w")

        tk.Label(form, text="End Date (YYYY-MM-DD):").grid(row=4, column=0, sticky="w")
        self.end_var = tk.StringVar(value=today.strftime("%Y-%m-%d"))
        tk.Entry(form, textvariable=self.end_var, width=35).grid(row=4, column=1, sticky="w")

        tk.Label(form, text="Target Folder:").grid(row=5, column=0, sticky="w")
        folder_frame = tk.Frame(form)
        folder_frame.grid(row=5, column=1, sticky="w")
        self.folder_label = tk.Label(folder_frame, text="(not selected)", fg="gray", width=25, anchor="w")
        self.folder_label.pack(side=tk.LEFT)
        tk.Button(folder_frame, text="Browse...", command=self._choose_folder).pack(side=tk.LEFT, padx=5)

        self.include_network_var = tk.BooleanVar(value=True)
        tk.Checkbutton(form, text="Include Network totals (BSS tab)",
                        variable=self.include_network_var).grid(row=6, column=0, columnspan=2, sticky="w", pady=(6, 0))

        self.run_btn = tk.Button(self.root, text="Fetch & Generate Report", font=("Arial", 11, "bold"),
                                  bg="blue", fg="white", command=self._start_run)
        self.run_btn.pack(pady=10)

        self.log_text = tk.Text(self.root, height=16, width=68, state=tk.DISABLED, bg="#f5f5f5")
        self.log_text.pack(padx=10, pady=(0, 10), fill=tk.BOTH, expand=True)

    def _choose_folder(self):
        d = filedialog.askdirectory(title="Select the Target Base Folder")
        if d:
            self.target_dir = d
            self.folder_label.config(text=d, fg="black")

    def _log(self, msg):
        self.log_queue.put(msg)

    def _poll_log(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_text.config(state=tk.NORMAL)
                self.log_text.insert(tk.END, msg + "\n")
                self.log_text.see(tk.END)
                self.log_text.config(state=tk.DISABLED)
        except queue.Empty:
            pass
        self.root.after(200, self._poll_log)

    def _start_run(self):
        email = self.email_var.get().strip()
        password = self.password_var.get()
        region = self.region_var.get()
        start = self.start_var.get().strip()
        end = self.end_var.get().strip()

        if not email or not password:
            messagebox.showerror("Missing Info", "Enter your username and password.")
            return
        if not self.target_dir:
            messagebox.showerror("Missing Info", "Choose a target folder.")
            return

        self.run_btn.config(state=tk.DISABLED, text="Working...")
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state=tk.DISABLED)

        thread = threading.Thread(
            target=self._run_worker,
            args=(email, password, region, start, end, self.target_dir, self.include_network_var.get()),
            daemon=True,
        )
        thread.start()

    def _run_worker(self, email, password, region, start, end, target_dir, include_network):
        try:
            session = requests.Session()
            self._log("Logging in...")
            user = kpi_fetch.login(session, email, password)
            self._log(f"Logged in as {user['FirstName']} {user['LastName']}")

            kpi_fetch.fetch_region_report(session, region, start, end, target_dir, include_network, self._log)

            self._log("Building Excel report...")
            output_path = core.build_excel_report(target_dir, region)
            self._log(f"Done! Report saved to:\n{output_path}")
            self.root.after(0, lambda: messagebox.showinfo("Success", f"Report generated:\n{output_path}"))
        except Exception as e:
            self._log(f"ERROR: {e}")
            self.root.after(0, lambda: messagebox.showerror("Failed", str(e)))
        finally:
            self.root.after(0, lambda: self.run_btn.config(state=tk.NORMAL, text="Fetch & Generate Report"))


if __name__ == "__main__":
    root = tk.Tk()
    app = KpiAutoApp(root)
    root.mainloop()
