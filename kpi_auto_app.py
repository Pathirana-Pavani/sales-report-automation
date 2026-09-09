import os
import queue
import threading
from datetime import datetime, timedelta

import requests
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

import report_builder as core
import kpi_fetch
import pptx_report


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
            excel_path = core.build_excel_report(target_dir, region)
            self._log(f"Excel report saved to:\n{excel_path}")

            templates_dir = os.path.join(os.path.dirname(__file__), "templates")
            template_path = os.path.join(templates_dir, f"{region}.pptx")
            excel_paths_for_pptx = [excel_path]

            if not os.path.exists(template_path):
                combo_name = next(
                    (name for name, members in kpi_fetch.COMBO_PRESENTATION_GROUPS.items() if region in members),
                    None,
                )
                if combo_name:
                    template_path = os.path.join(templates_dir, f"{combo_name}.pptx")
                    sibling_regions = [r for r in kpi_fetch.COMBO_PRESENTATION_GROUPS[combo_name] if r != region]
                    for sibling in sibling_regions:
                        self._log(f"Also fetching {sibling} (shares this presentation)...")
                        kpi_fetch.fetch_region_report(session, sibling, start, end, target_dir, include_network, self._log)
                        excel_paths_for_pptx.append(core.build_excel_report(target_dir, sibling))

            pptx_path = None
            if os.path.exists(template_path):
                self._log("Updating presentation charts...")
                label = os.path.splitext(os.path.basename(template_path))[0]
                pptx_path = os.path.join(os.path.dirname(excel_path), f"{label}_Monthly_Review.pptx")
                pptx_report.build_pptx_report(template_path, excel_paths_for_pptx, pptx_path, log=self._log)
                self._log(f"Presentation saved to:\n{pptx_path}")
            else:
                self._log(f"No presentation template configured for {region} yet — skipping.")

            msg = f"Excel report:\n{excel_path}"
            if pptx_path:
                msg += f"\n\nPresentation:\n{pptx_path}"
            self.root.after(0, lambda: messagebox.showinfo("Success", msg))
        except Exception as e:
            self._log(f"ERROR: {e}")
            self.root.after(0, lambda: messagebox.showerror("Failed", str(e)))
        finally:
            self.root.after(0, lambda: self.run_btn.config(state=tk.NORMAL, text="Fetch & Generate Report"))


if __name__ == "__main__":
    root = tk.Tk()
    app = KpiAutoApp(root)
    root.mainloop()
