import os
import shutil
import time
import queue
import glob
import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from watchdog.observers.polling import PollingObserver as Observer
from watchdog.events import FileSystemEventHandler

# --- CONFIGURATION ---
DISTRICTS = {
    "Central": ["Badulla", "Kandy", "Matale", "Nuwara Eliya", "Network"],
    "South": ["Galle", "Matara", "Hambantota", "Monaragala", "Ratnapura", "Network"], 
    "North East": ["Jaffna", "Kilinochchi", "Mannar", "Mullaitivu", "Vavuniya", "Ampara", "Batticaloa", "Trincomalee", "Network"]
}

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

class _ReportContext:
    """Minimal stand-in for AppUI carrying just the state the sheet-building
    methods need (target_base_dir, selected_region, regional_dfs), so the
    report can be built without a Tkinter/watchdog session — reused by both
    the manual app and headless/automated callers."""
    def __init__(self, target_base_dir, region):
        self.target_base_dir = target_base_dir
        self.selected_region = region
        self.regional_dfs = {}


def build_excel_report(target_base_dir, region):
    ctx = _ReportContext(target_base_dir, region)
    base_folder = os.path.join(target_base_dir, "Sales Report", region)
    output_path = os.path.join(base_folder, f"{region}_Region_Final_Sales_Report.xlsx")

    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    workbook = writer.book
    date_fmt = workbook.add_format({'num_format': 'm/d/yyyy'})

    if region == "Central":
        AppUI.process_2g_voice_sheet(ctx, base_folder, writer, workbook)
        AppUI.process_3g_voice_sheet(ctx, base_folder, writer, workbook)
        AppUI.process_3g_data_sheet(ctx, base_folder, writer, workbook)
        AppUI.process_4g_voice_sheet(ctx, base_folder, writer, workbook)
        AppUI.process_4g_data_sheet(ctx, base_folder, writer, workbook)

    elif region in ["South", "North East"]:
        for tech, folder, traffic_logic, unit in [
            ('2G Voice', '2G Voice', lambda df: df['TCH HR (Erl)'] + df['TCH FR (Erl)'], 'Erlang'),
            ('3G Voice', '3G Voice', lambda df: df['CS AMR (Erl)'], 'Erlang'),
            ('3G Data', '3G Data', lambda df: df['PS Traffic Volume (GByte)'], 'TB'),
            ('4G Data', '4G Data', lambda df: df['Traffic Volume (GByte)'], 'TB')
        ]:
            files = glob.glob(os.path.join(base_folder, folder, "*.csv"))
            if files:
                dfs = {}
                for f in files:
                    loc = os.path.basename(f).split('_')[-1].replace('.csv', '')
                    df = pd.read_csv(f)
                    df['Date'] = pd.to_datetime(df['DateTime'], format='mixed')
                    dfs[loc] = pd.DataFrame({'Date': df['Date'], 'Traffic': traffic_logic(df)})
                metric_name = 'Voice Traffic' if 'Voice' in tech else 'Data Traffic'
                AppUI._merge_and_chart_south_north(ctx, dfs, writer, workbook, tech, metric_name, unit, date_fmt)

        AppUI.process_split_volte_sheets(ctx, base_folder, writer, workbook)
        AppUI.process_total_data_sheet(ctx, base_folder, writer, workbook)
        AppUI.process_total_voice_sheet(ctx, base_folder, writer, workbook)

    writer.close()
    return output_path


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

    # ==========================================
    # CENTRAL SHEET PROCESSING
    # ==========================================
    def process_2g_voice_sheet(self, base_folder, writer, workbook):
        csv_files = glob.glob(os.path.join(base_folder, '2G Voice', "*.csv"))
        if not csv_files: return
        data_frames = {}
        for file in csv_files:
            loc = os.path.basename(file).split('_')[-1].replace('.csv', '')
            df = pd.read_csv(file)
            data_frames[loc] = pd.DataFrame({'Date': pd.to_datetime(df['DateTime']).dt.strftime('%m/%d/%Y'), 'Traffic': df['TCH HR (Erl)'] + df['TCH FR (Erl)']})
        self._merge_and_chart(data_frames, writer, workbook, '2G Voice', 'Erlang')

    def process_3g_voice_sheet(self, base_folder, writer, workbook):
        csv_files = glob.glob(os.path.join(base_folder, '3G Voice', "*.csv"))
        if not csv_files: return
        data_frames = {}
        for file in csv_files:
            loc = os.path.basename(file).split('_')[-1].replace('.csv', '')
            df = pd.read_csv(file)
            data_frames[loc] = pd.DataFrame({'Date': pd.to_datetime(df['DateTime']).dt.strftime('%m/%d/%Y'), 'Traffic': df['CS AMR (Erl)']})
        self._merge_and_chart(data_frames, writer, workbook, '3G Voice', 'Erlang')
        
    def process_3g_data_sheet(self, base_folder, writer, workbook):
        csv_files = glob.glob(os.path.join(base_folder, '3G Data', "*.csv"))
        if not csv_files: return
        data_frames = {}
        for file in csv_files:
            loc = os.path.basename(file).split('_')[-1].replace('.csv', '')
            df = pd.read_csv(file)
            data_frames[loc] = pd.DataFrame({'Date': pd.to_datetime(df['DateTime']).dt.strftime('%m/%d/%Y'), 'Traffic': df['PS Traffic Volume (GByte)'] / 1024})
        self._merge_and_chart(data_frames, writer, workbook, '3G Data', 'TB')

    def process_4g_voice_sheet(self, base_folder, writer, workbook):
        csv_files = glob.glob(os.path.join(base_folder, '4G Voice', "*.csv"))
        if not csv_files: return
        data_frames = {}
        for file in csv_files:
            loc = os.path.basename(file).split('_')[-1].replace('.csv', '')
            df = pd.read_csv(file)
            data_frames[loc] = pd.DataFrame({'Date': pd.to_datetime(df['DateTime']).dt.strftime('%m/%d/%Y'), 'Traffic': df['VoLTE Traffic (Erl)']})
        self._merge_and_chart(data_frames, writer, workbook, '4G Voice', 'Erlang')

    def process_4g_data_sheet(self, base_folder, writer, workbook):
        csv_files = glob.glob(os.path.join(base_folder, '4G Data', "*.csv"))
        if not csv_files: return
        data_frames = {}
        for file in csv_files:
            loc = os.path.basename(file).split('_')[-1].replace('.csv', '')
            df = pd.read_csv(file)
            data_frames[loc] = pd.DataFrame({'Date': pd.to_datetime(df['DateTime']).dt.strftime('%m/%d/%Y'), 'Traffic': df['Traffic Volume (GByte)'] / 1024})
        self._merge_and_chart(data_frames, writer, workbook, '4G Data', 'TB')

    def _merge_and_chart(self, data_frames, writer, workbook, sheet_name, unit):
        master_df = None
        for loc, df in data_frames.items():
            df = df.rename(columns={'Traffic': loc.title()})
            if master_df is None: master_df = df
            else: master_df = pd.merge(master_df, df, on='Date', how='outer')
        if 'Nuwaraeliya' in master_df.columns: master_df.rename(columns={'Nuwaraeliya': 'Nuwara Eliya'}, inplace=True)
        master_df = master_df.fillna(0)
        districts = [d for d in DISTRICTS[self.selected_region] if d != "Network"]
        available = [d for d in districts if d in master_df.columns]
        master_df['Cluster'] = master_df[available].sum(axis=1)
        final_df = master_df[['Date'] + [c for c in available if c in master_df.columns] + ['Network', 'Cluster']]
        final_df.to_excel(writer, sheet_name=sheet_name, index=False)
        worksheet = writer.sheets[sheet_name]; max_row = len(final_df); headers = list(final_df.columns); date_col = headers.index('Date')
        if 'Network' not in headers or 'Cluster' not in headers: return
        net_col, cluster_col = headers.index('Network'), headers.index('Cluster')
        c1_prim = workbook.add_chart({'type': 'line'})
        for dist in available: c1_prim.add_series({'name': [sheet_name, 0, headers.index(dist)], 'categories': [sheet_name, 1, date_col, max_row, date_col], 'values': [sheet_name, 1, headers.index(dist), max_row, headers.index(dist)]})
        c1_prim.set_title({'name': f'{sheet_name} ({unit}): Districts Vs. Network'}); c1_prim.set_legend({'position': 'bottom'}); 
        c1_prim.set_x_axis({'name': 'Date'}) 
        c1_prim.set_y_axis({'name': f'Districts ({unit})'})
        c1_sec = workbook.add_chart({'type': 'line'}); c1_sec.add_series({'name': [sheet_name, 0, net_col], 'categories': [sheet_name, 1, date_col, max_row, date_col], 'values': [sheet_name, 1, net_col, max_row, net_col], 'y2_axis': True}); c1_prim.combine(c1_sec); c1_sec.set_y2_axis({'name': 'Network Traffic'}); worksheet.insert_chart('J2', c1_prim, {'x_scale': 1.5, 'y_scale': 1.2})
        c2_prim = workbook.add_chart({'type': 'line'}); c2_prim.add_series({'name': [sheet_name, 0, cluster_col], 'categories': [sheet_name, 1, date_col, max_row, date_col], 'values': [sheet_name, 1, cluster_col, max_row, cluster_col]}); c2_prim.set_title({'name': f'{sheet_name} ({unit}) – {self.selected_region} Cluster Vs Network'}); c2_prim.set_legend({'position': 'bottom'}); 
        c2_prim.set_x_axis({'name': 'Date'})
        c2_prim.set_y_axis({'name': f'Cluster ({unit})'})
        c2_sec = workbook.add_chart({'type': 'line'}); c2_sec.add_series({'name': [sheet_name, 0, net_col], 'categories': [sheet_name, 1, date_col, max_row, date_col], 'values': [sheet_name, 1, net_col, max_row, net_col], 'y2_axis': True}); c2_prim.combine(c2_sec); c2_sec.set_y2_axis({'name': 'Network Traffic'}); worksheet.insert_chart('J20', c2_prim, {'x_scale': 1.5, 'y_scale': 1.2})

    # ==========================================
    # CUSTOM SHEETS (South / North East)
    # ==========================================
    def _merge_and_chart_south_north(self, data_frames, writer, workbook, sheet_name, metric_name, unit, date_fmt):
        master_df = None
        for loc, df in data_frames.items():
            df = df.rename(columns={'Traffic': loc.title()})
            if master_df is None: master_df = df
            else: master_df = pd.merge(master_df, df, on='Date', how='outer')
            
        master_df['Date'] = pd.to_datetime(master_df['Date'])
        master_df = master_df.sort_values(by='Date').fillna(0)
        
        districts = [d for d in DISTRICTS[self.selected_region] if d != "Network"]
        available = [d for d in districts if d in master_df.columns]
        
        df_for_memory = master_df[['Date'] + available + ['Network']] if 'Network' in master_df.columns else master_df[['Date'] + available]
        
        if not hasattr(self, 'regional_dfs'): self.regional_dfs = {}
        self.regional_dfs[sheet_name] = df_for_memory.copy()
        
        final_df = df_for_memory.copy()
        final_df['Date'] = final_df['Date'].dt.month.astype(str) + '/' + final_df['Date'].dt.day.astype(str) + '/' + final_df['Date'].dt.year.astype(str)
        
        final_df.to_excel(writer, sheet_name=sheet_name, index=False)
        worksheet = writer.sheets[sheet_name]
        worksheet.set_column('A:A', 15, date_fmt)
        
        max_row = len(final_df)
        headers = list(final_df.columns)
        date_col = headers.index('Date')

        if 'Network' not in headers: return
        net_col = headers.index('Network')
        adj = "NorthEast" if self.selected_region == "North East" else "Southern"
        
        c1_prim = workbook.add_chart({'type': 'line'})
        for dist in available: 
            c1_prim.add_series({'name': [sheet_name, 0, headers.index(dist)], 'categories': [sheet_name, 1, date_col, max_row, date_col], 'values': [sheet_name, 1, headers.index(dist), max_row, headers.index(dist)]})
        c1_prim.set_title({'name': f'Data and Voice Traffic Behavior of {adj} Cluster – {sheet_name} Traffic'})
        c1_prim.set_legend({'position': 'bottom'})
        c1_prim.set_x_axis({'name': 'Date', 'date_axis': True}) 
        c1_prim.set_y_axis({'name': f'{metric_name} {adj} Region ({unit})'})
        
        c1_sec = workbook.add_chart({'type': 'line'})
        c1_sec.add_series({'name': [sheet_name, 0, net_col], 'categories': [sheet_name, 1, date_col, max_row, date_col], 'values': [sheet_name, 1, net_col, max_row, net_col], 'y2_axis': True})
        c1_sec.set_y2_axis({'name': f'{metric_name} Network ({unit})'})
        c1_prim.combine(c1_sec)
        worksheet.insert_chart('J2', c1_prim, {'x_scale': 1.5, 'y_scale': 1.2})

    def process_split_volte_sheets(self, base_folder, writer, workbook):
        csv_files = glob.glob(os.path.join(base_folder, '4G Voice', "*.csv"))
        if not csv_files: return
        
        data_frames = {}
        for file in csv_files:
            loc = os.path.basename(file).split('_')[-1].replace('.csv', '')
            df = pd.read_csv(file)
            df['Date'] = pd.to_datetime(df['DateTime'], format='mixed')
            data_frames[loc] = pd.DataFrame({'Date': df['Date'], 'Traffic': df['VoLTE Traffic (Erl)']})
            
        master_df = None
        for loc, df in data_frames.items():
            df = df.rename(columns={'Traffic': loc.title()})
            if master_df is None: master_df = df
            else: master_df = pd.merge(master_df, df, on='Date', how='outer')
            
        master_df['Date'] = pd.to_datetime(master_df['Date'])
        master_df = master_df.sort_values(by='Date').fillna(0)
        
        if not hasattr(self, 'regional_dfs'): self.regional_dfs = {}
        self.regional_dfs['VoLTE'] = master_df.copy()

        districts = [d for d in DISTRICTS[self.selected_region] if d != "Network"]
        available = [d for d in districts if d in master_df.columns]
        adj = "NorthEast" if self.selected_region == "North East" else "Southern"
        date_fmt = workbook.add_format({'num_format': 'm/d/yyyy'})

        # --- SHEET 1: 4G Voice - Cluster ---
        cols_c = ['Date'] + available
        if self.selected_region == "North East" and 'Network' in master_df.columns:
            cols_c.append('Network')
            
        cluster_df = master_df[cols_c].copy()
        cluster_df['Date'] = cluster_df['Date'].dt.month.astype(str) + '/' + cluster_df['Date'].dt.day.astype(str) + '/' + cluster_df['Date'].dt.year.astype(str)
        
        sheet_name_c = '4G Voice - Cluster'
        cluster_df.to_excel(writer, sheet_name=sheet_name_c, index=False)
        ws_c = writer.sheets[sheet_name_c]
        ws_c.set_column('A:A', 15, date_fmt)
        max_row = len(cluster_df)
        
        chart_c = workbook.add_chart({'type': 'line'})
        for dist in available:
            col_idx = list(cluster_df.columns).index(dist)
            chart_c.add_series({'name': [sheet_name_c, 0, col_idx], 'categories': [sheet_name_c, 1, 0, max_row, 0], 'values': [sheet_name_c, 1, col_idx, max_row, col_idx]})
        chart_c.set_title({'name': f'VoLTE traffic Behavior of {adj} Cluster'})
        chart_c.set_y_axis({'name': 'VoLTE traffic (Erl)'})
        chart_c.set_x_axis({'name': 'Date', 'date_axis': True}) 
        chart_c.set_legend({'position': 'bottom'})
        ws_c.insert_chart('J2', chart_c, {'x_scale': 1.5, 'y_scale': 1.2})

        # --- SHEET 2: 4G Voice - Network ---
        if 'Network' in master_df.columns:
            net_df = master_df[['Date', 'Network']].copy()
            net_df['Date'] = net_df['Date'].dt.month.astype(str) + '/' + net_df['Date'].dt.day.astype(str) + '/' + net_df['Date'].dt.year.astype(str)
            
            sheet_name_n = '4G Voice - Network'
            net_df.to_excel(writer, sheet_name=sheet_name_n, index=False)
            ws_n = writer.sheets[sheet_name_n]
            ws_n.set_column('A:A', 15, date_fmt)
            
            chart_n = workbook.add_chart({'type': 'line'})
            chart_n.add_series({'name': [sheet_name_n, 0, 1], 'categories': [sheet_name_n, 1, 0, max_row, 0], 'values': [sheet_name_n, 1, 1, max_row, 1]})
            chart_n.set_title({'name': 'VoLTE traffic Behavior of the Network'})
            chart_n.set_y_axis({'name': 'VoLTE traffic (Erl)'})
            chart_n.set_x_axis({'name': 'Date', 'date_axis': True}) 
            chart_n.set_legend({'position': 'bottom'})
            ws_n.insert_chart('D2', chart_n, {'x_scale': 1.5, 'y_scale': 1.2})

    def process_total_data_sheet(self, base_folder, writer, workbook):
        if not hasattr(self, 'regional_dfs') or '3G Data' not in self.regional_dfs or '4G Data' not in self.regional_dfs: return
        
        df3 = self.regional_dfs['3G Data'].set_index('Date')
        df4 = self.regional_dfs['4G Data'].set_index('Date')
        total_df = df3.add(df4, fill_value=0).reset_index()
        
        districts = [d for d in DISTRICTS[self.selected_region] if d != "Network"]
        available = [d for d in districts if d in total_df.columns]
        
        if self.selected_region == "North East":
            total_df['Total North East Data Traffic'] = total_df[available].sum(axis=1)
            if 'Network' in total_df.columns:
                total_df = total_df.rename(columns={'Network': 'Total Network Data Traffic'})
            cols = ['Date'] + available
            if 'Total Network Data Traffic' in total_df.columns: cols.append('Total Network Data Traffic')
            cols.append('Total North East Data Traffic')
            final_df = total_df[cols].copy()
        else:
            final_df = total_df[['Date'] + available].copy()
            
        final_df['Date'] = final_df['Date'].dt.month.astype(str) + '/' + final_df['Date'].dt.day.astype(str) + '/' + final_df['Date'].dt.year.astype(str)
        
        sheet_name = 'Total Data'
        final_df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]
        
        date_fmt = workbook.add_format({'num_format': 'm/d/yyyy'})
        ws.set_column('A:A', 15, date_fmt)
        max_row = len(final_df)
        
        c1_prim = workbook.add_chart({'type': 'line'})
        
        if self.selected_region == "North East":
            if 'Total North East Data Traffic' in final_df.columns:
                c_idx = list(final_df.columns).index('Total North East Data Traffic')
                c1_prim.add_series({'name': [sheet_name, 0, c_idx], 'categories': [sheet_name, 1, 0, max_row, 0], 'values': [sheet_name, 1, c_idx, max_row, c_idx]})
            
            c1_prim.set_title({'name': 'Data and Voice Traffic Behavior of NorthEast Cluster – Total Data Traffic (3G/4G)'})
            c1_prim.set_y_axis({'name': 'Total NorthEast Cluster Data Traffic (TB)'})
            c1_prim.set_x_axis({'name': 'Date', 'date_axis': True}) 
            c1_prim.set_legend({'position': 'bottom'})

            # Add Network as Secondary Axis for Combo Graph
            if 'Total Network Data Traffic' in final_df.columns:
                c1_sec = workbook.add_chart({'type': 'line'})
                n_idx = list(final_df.columns).index('Total Network Data Traffic')
                c1_sec.add_series({'name': [sheet_name, 0, n_idx], 'categories': [sheet_name, 1, 0, max_row, 0], 'values': [sheet_name, 1, n_idx, max_row, n_idx], 'y2_axis': True})
                c1_sec.set_y2_axis({'name': 'Total Network Data Traffic (TB)'})
                c1_prim.combine(c1_sec)
                
            ws.insert_chart('J2', c1_prim, {'x_scale': 1.5, 'y_scale': 1.2})
        else:
            adj = "Southern"
            for dist in available:
                col_idx = list(final_df.columns).index(dist)
                c1_prim.add_series({'name': [sheet_name, 0, col_idx], 'categories': [sheet_name, 1, 0, max_row, 0], 'values': [sheet_name, 1, col_idx, max_row, col_idx]})
            c1_prim.set_title({'name': f'Data and Voice Traffic Behavior of {adj} Cluster – Total Data Traffic'})
            c1_prim.set_y_axis({'name': 'Total Data Traffic (TB)'})
            c1_prim.set_x_axis({'name': 'Date', 'date_axis': True}) 
            c1_prim.set_legend({'position': 'bottom'})
            ws.insert_chart('J2', c1_prim, {'x_scale': 1.5, 'y_scale': 1.2})

    def process_total_voice_sheet(self, base_folder, writer, workbook):
        if not hasattr(self, 'regional_dfs') or '2G Voice' not in self.regional_dfs or '3G Voice' not in self.regional_dfs or 'VoLTE' not in self.regional_dfs: return
        
        df2 = self.regional_dfs['2G Voice'].set_index('Date')
        df3 = self.regional_dfs['3G Voice'].set_index('Date')
        df4 = self.regional_dfs['VoLTE'].set_index('Date')
        
        # Add 4G Voice (VoLTE) back into the Total Voice calculation
        total_df = df2.add(df3, fill_value=0).add(df4, fill_value=0).reset_index()
        
        districts = [d for d in DISTRICTS[self.selected_region] if d != "Network"]
        available = [d for d in districts if d in total_df.columns]
        
        if self.selected_region == "North East":
            total_df['Total North East Voice Traffic'] = total_df[available].sum(axis=1)
            if 'Network' in total_df.columns:
                total_df = total_df.rename(columns={'Network': 'Total Network Voiced Traffic'})
            cols = ['Date'] + available
            if 'Total Network Voiced Traffic' in total_df.columns: cols.append('Total Network Voiced Traffic')
            cols.append('Total North East Voice Traffic')
            final_df = total_df[cols].copy()
        else:
            final_df = total_df[['Date'] + available].copy()
            
        final_df['Date'] = final_df['Date'].dt.month.astype(str) + '/' + final_df['Date'].dt.day.astype(str) + '/' + final_df['Date'].dt.year.astype(str)
        
        sheet_name = 'Total Voice'
        final_df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]
        
        date_fmt = workbook.add_format({'num_format': 'm/d/yyyy'})
        ws.set_column('A:A', 15, date_fmt)
        max_row = len(final_df)
        
        c1_prim = workbook.add_chart({'type': 'line'})
        
        if self.selected_region == "North East":
            if 'Total North East Voice Traffic' in final_df.columns:
                c_idx = list(final_df.columns).index('Total North East Voice Traffic')
                c1_prim.add_series({'name': [sheet_name, 0, c_idx], 'categories': [sheet_name, 1, 0, max_row, 0], 'values': [sheet_name, 1, c_idx, max_row, c_idx]})
            
            c1_prim.set_title({'name': 'Data and Voice Traffic Behavior of NorthEast Cluster – Total Voice Traffic (2G/3G/4G)'})
            c1_prim.set_y_axis({'name': 'Total NorthEast Cluster Voice Traffic (Erlang)'})
            c1_prim.set_x_axis({'name': 'Date', 'date_axis': True}) 
            c1_prim.set_legend({'position': 'bottom'})

            # Add Network as Secondary Axis for Combo Graph
            if 'Total Network Voiced Traffic' in final_df.columns:
                c1_sec = workbook.add_chart({'type': 'line'})
                n_idx = list(final_df.columns).index('Total Network Voiced Traffic')
                c1_sec.add_series({'name': [sheet_name, 0, n_idx], 'categories': [sheet_name, 1, 0, max_row, 0], 'values': [sheet_name, 1, n_idx, max_row, n_idx], 'y2_axis': True})
                c1_sec.set_y2_axis({'name': 'Total Network Voice Traffic (Erlang)'})
                c1_prim.combine(c1_sec)
                
            ws.insert_chart('J2', c1_prim, {'x_scale': 1.5, 'y_scale': 1.2})
        else:
            adj = "Southern"
            for dist in available:
                col_idx = list(final_df.columns).index(dist)
                c1_prim.add_series({'name': [sheet_name, 0, col_idx], 'categories': [sheet_name, 1, 0, max_row, 0], 'values': [sheet_name, 1, col_idx, max_row, col_idx]})
            c1_prim.set_title({'name': f'Data and Voice Traffic Behavior of {adj} Cluster – Total Voice Traffic'})
            c1_prim.set_y_axis({'name': 'Total Voice Traffic (Erlang)'})
            c1_prim.set_x_axis({'name': 'Date', 'date_axis': True}) 
            c1_prim.set_legend({'position': 'bottom'})
            ws.insert_chart('J2', c1_prim, {'x_scale': 1.5, 'y_scale': 1.2})

    def close_app(self):
        try: self.observer.stop()
        except: pass
        self.root.destroy(); os._exit(0)

# process_2g/3g/4g_voice_sheet etc. call self._merge_and_chart(...) internally,
# so _ReportContext needs it bound too, not just the methods build_excel_report
# calls directly.
_ReportContext._merge_and_chart = AppUI._merge_and_chart

if __name__ == "__main__":
    root = tk.Tk()
    app = AppUI(root)
    try: root.mainloop()
    except KeyboardInterrupt: app.close_app()