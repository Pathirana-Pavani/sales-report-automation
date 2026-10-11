"""Excel report generation from the per-location KPI CSVs — pure data/Excel
logic, no GUI or filesystem-watching dependencies, so it can run anywhere
(desktop app, Streamlit Cloud, a cron job, etc).
"""
import os
import glob

import pandas as pd

DISTRICTS = {
    "Central": ["Badulla", "Kandy", "Matale", "Nuwara Eliya", "Network"],
    "South": ["Galle", "Matara", "Hambantota", "Monaragala", "Ratnapura", "Network"],
    "North East": ["Jaffna", "Kilinochchi", "Mannar", "Mullaitivu", "Vavuniya", "Ampara", "Batticaloa", "Trincomalee", "Network"],
    "Gampaha": ["Gampaha", "Network"],
    "Kegalle-Puttalama": ["Kegalle", "Puttalam", "Network"],
    "Anuradhapura-Kurunegala": ["Anuradhapura", "Polonnaruwa", "Kurunegala", "Network"],
    "Colombo-Kaluthara": ["Colombo", "Kalutara", "Network"],
}

# For the non-Central sheet titles/charts ("Data and Voice Traffic Behavior
# of <ADJ> Cluster..."). South and North East keep their original wording;
# new regions default to their own name.
REGION_ADJECTIVE = {"South": "Southern", "North East": "NorthEast"}


class ReportBuilder:
    def __init__(self, target_base_dir, region, periods=None, workbook=None):
        self.target_base_dir = target_base_dir
        self.selected_region = region
        self.regional_dfs = {}
        # periods = {"pre": (start, end), "post": (start, end)} (date objects),
        # the same for every sheet in one workbook; None disables the tables.
        self.periods = periods
        self.workbook = workbook

    # ==========================================
    # PRE / POST TRAFFIC SUMMARY TABLE
    # ==========================================
    @staticmethod
    def _excel_col(idx):
        name = ""
        idx += 1
        while idx:
            idx, rem = divmod(idx - 1, 26)
            name = chr(65 + rem) + name
        return name

    def _row_ranges(self, dates, start, end):
        """1-based Excel row runs (first, last) whose date is within [start, end]."""
        rows = [i + 2 for i, d in enumerate(dates) if start <= d <= end]
        runs = []
        for r in rows:
            if runs and r == runs[-1][1] + 1: runs[-1][1] = r
            else: runs.append([r, r])
        return runs

    def _add_summary_table(self, worksheet, sheet_name, final_df, unit, row, col):
        """Writes a Pre / Post / Improvement table (average traffic per column)
        as live Excel formulas referencing the sheet's own data."""
        if not self.periods or self.workbook is None: return
        headers = list(final_df.columns)
        value_cols = [h for h in headers if h != 'Date']
        # Network columns last, like the reference report (districts, Cluster, Network)
        value_cols = [h for h in value_cols if 'Network' not in h] + [h for h in value_cols if 'Network' in h]
        dates = pd.to_datetime(final_df['Date'], format='%m/%d/%Y').dt.date.tolist()

        wb = self.workbook
        head_fmt = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#D9E1F2', 'align': 'center'})
        label_fmt = wb.add_format({'bold': True, 'border': 1})
        num_fmt = wb.add_format({'border': 1, 'num_format': '#,##0.00'})
        pct_fmt = wb.add_format({'border': 1, 'num_format': '0.00%', 'bold': True})

        def fmt_range(rng): return f"{rng[0].strftime('%b %d, %Y')} - {rng[1].strftime('%b %d, %Y')}"
        pre, post = self.periods['pre'], self.periods['post']
        worksheet.write(row, col, f"{sheet_name} ({unit})", head_fmt)
        worksheet.write(row + 1, col, f"Pre Traffic ({fmt_range(pre)})", label_fmt)
        worksheet.write(row + 2, col, f"Post Traffic ({fmt_range(post)})", label_fmt)
        worksheet.write(row + 3, col, "Improvement", label_fmt)
        worksheet.set_column(col, col, 38)

        pre_runs = self._row_ranges(dates, *pre)
        post_runs = self._row_ranges(dates, *post)
        for i, name in enumerate(value_cols):
            c = col + 1 + i
            letter = self._excel_col(headers.index(name))
            worksheet.write(row, c, name, head_fmt)
            avgs = []
            for r_off, runs in ((1, pre_runs), (2, post_runs)):
                if runs:
                    ref = ",".join(f"{letter}{a}:{letter}{b}" for a, b in runs)
                    # cached value so viewers that don't recalculate still show numbers
                    vals = pd.concat([final_df[name].iloc[a - 2:b - 1] for a, b in runs])
                    avg = float(vals.mean())
                    worksheet.write_formula(row + r_off, c, f"=AVERAGE({ref})", num_fmt, avg)
                    avgs.append(avg)
                else:
                    worksheet.write(row + r_off, c, "n/a", num_fmt)
                    avgs.append(None)
            imp = (avgs[1] - avgs[0]) / avgs[0] if None not in avgs and avgs[0] else "n/a"
            pre_c = f"{self._excel_col(c)}{row + 2}"
            post_c = f"{self._excel_col(c)}{row + 3}"
            worksheet.write_formula(row + 3, c, f'=IFERROR(({post_c}-{pre_c})/{pre_c},"n/a")', pct_fmt, imp)

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
        self._add_summary_table(worksheet, sheet_name, final_df, unit, 38, 9)
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

        self.regional_dfs[sheet_name] = df_for_memory.copy()

        final_df = df_for_memory.copy()
        final_df['Date'] = final_df['Date'].dt.month.astype(str) + '/' + final_df['Date'].dt.day.astype(str) + '/' + final_df['Date'].dt.year.astype(str)

        final_df.to_excel(writer, sheet_name=sheet_name, index=False)
        worksheet = writer.sheets[sheet_name]
        worksheet.set_column('A:A', 15, date_fmt)
        self._add_summary_table(worksheet, sheet_name, final_df, unit, 20, 9)

        max_row = len(final_df)
        headers = list(final_df.columns)
        date_col = headers.index('Date')

        if 'Network' not in headers: return
        net_col = headers.index('Network')
        adj = REGION_ADJECTIVE.get(self.selected_region, self.selected_region)

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

        self.regional_dfs['VoLTE'] = master_df.copy()

        districts = [d for d in DISTRICTS[self.selected_region] if d != "Network"]
        available = [d for d in districts if d in master_df.columns]
        adj = REGION_ADJECTIVE.get(self.selected_region, self.selected_region)
        date_fmt = workbook.add_format({'num_format': 'm/d/yyyy'})

        # --- SHEET 1: 4G Voice - Cluster ---
        cols_c = ['Date'] + available
        if self.selected_region != "South" and 'Network' in master_df.columns:
            cols_c.append('Network')

        cluster_df = master_df[cols_c].copy()
        cluster_df['Date'] = cluster_df['Date'].dt.month.astype(str) + '/' + cluster_df['Date'].dt.day.astype(str) + '/' + cluster_df['Date'].dt.year.astype(str)

        sheet_name_c = '4G Voice - Cluster'
        cluster_df.to_excel(writer, sheet_name=sheet_name_c, index=False)
        ws_c = writer.sheets[sheet_name_c]
        ws_c.set_column('A:A', 15, date_fmt)
        self._add_summary_table(ws_c, sheet_name_c, cluster_df, 'Erlang', 20, 9)
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
            self._add_summary_table(ws_n, sheet_name_n, net_df, 'Erlang', 20, 3)

            chart_n = workbook.add_chart({'type': 'line'})
            chart_n.add_series({'name': [sheet_name_n, 0, 1], 'categories': [sheet_name_n, 1, 0, max_row, 0], 'values': [sheet_name_n, 1, 1, max_row, 1]})
            chart_n.set_title({'name': 'VoLTE traffic Behavior of the Network'})
            chart_n.set_y_axis({'name': 'VoLTE traffic (Erl)'})
            chart_n.set_x_axis({'name': 'Date', 'date_axis': True})
            chart_n.set_legend({'position': 'bottom'})
            ws_n.insert_chart('D2', chart_n, {'x_scale': 1.5, 'y_scale': 1.2})

    def process_total_data_sheet(self, base_folder, writer, workbook):
        if '3G Data' not in self.regional_dfs or '4G Data' not in self.regional_dfs: return

        df3 = self.regional_dfs['3G Data'].set_index('Date')
        df4 = self.regional_dfs['4G Data'].set_index('Date')
        total_df = df3.add(df4, fill_value=0).reset_index()

        districts = [d for d in DISTRICTS[self.selected_region] if d != "Network"]
        available = [d for d in districts if d in total_df.columns]

        adj = REGION_ADJECTIVE.get(self.selected_region, self.selected_region)
        total_col = f'Total {self.selected_region} Data Traffic'

        if self.selected_region != "South":
            total_df[total_col] = total_df[available].sum(axis=1)
            if 'Network' in total_df.columns:
                total_df = total_df.rename(columns={'Network': 'Total Network Data Traffic'})
            cols = ['Date'] + available
            if 'Total Network Data Traffic' in total_df.columns: cols.append('Total Network Data Traffic')
            cols.append(total_col)
            final_df = total_df[cols].copy()
        else:
            final_df = total_df[['Date'] + available].copy()

        final_df['Date'] = final_df['Date'].dt.month.astype(str) + '/' + final_df['Date'].dt.day.astype(str) + '/' + final_df['Date'].dt.year.astype(str)

        sheet_name = 'Total Data'
        final_df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]

        date_fmt = workbook.add_format({'num_format': 'm/d/yyyy'})
        ws.set_column('A:A', 15, date_fmt)
        self._add_summary_table(ws, sheet_name, final_df, 'TB', 20, 9)
        max_row = len(final_df)

        c1_prim = workbook.add_chart({'type': 'line'})

        if self.selected_region != "South":
            if total_col in final_df.columns:
                c_idx = list(final_df.columns).index(total_col)
                c1_prim.add_series({'name': [sheet_name, 0, c_idx], 'categories': [sheet_name, 1, 0, max_row, 0], 'values': [sheet_name, 1, c_idx, max_row, c_idx]})

            c1_prim.set_title({'name': f'Data and Voice Traffic Behavior of {adj} Cluster – Total Data Traffic (3G/4G)'})
            c1_prim.set_y_axis({'name': f'Total {adj} Cluster Data Traffic (TB)'})
            c1_prim.set_x_axis({'name': 'Date', 'date_axis': True})
            c1_prim.set_legend({'position': 'bottom'})

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
        if '2G Voice' not in self.regional_dfs or '3G Voice' not in self.regional_dfs or 'VoLTE' not in self.regional_dfs: return

        df2 = self.regional_dfs['2G Voice'].set_index('Date')
        df3 = self.regional_dfs['3G Voice'].set_index('Date')
        df4 = self.regional_dfs['VoLTE'].set_index('Date')

        total_df = df2.add(df3, fill_value=0).add(df4, fill_value=0).reset_index()

        districts = [d for d in DISTRICTS[self.selected_region] if d != "Network"]
        available = [d for d in districts if d in total_df.columns]

        adj = REGION_ADJECTIVE.get(self.selected_region, self.selected_region)
        total_col = f'Total {self.selected_region} Voice Traffic'

        if self.selected_region != "South":
            total_df[total_col] = total_df[available].sum(axis=1)
            if 'Network' in total_df.columns:
                total_df = total_df.rename(columns={'Network': 'Total Network Voiced Traffic'})
            cols = ['Date'] + available
            if 'Total Network Voiced Traffic' in total_df.columns: cols.append('Total Network Voiced Traffic')
            cols.append(total_col)
            final_df = total_df[cols].copy()
        else:
            final_df = total_df[['Date'] + available].copy()

        final_df['Date'] = final_df['Date'].dt.month.astype(str) + '/' + final_df['Date'].dt.day.astype(str) + '/' + final_df['Date'].dt.year.astype(str)

        sheet_name = 'Total Voice'
        final_df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]

        date_fmt = workbook.add_format({'num_format': 'm/d/yyyy'})
        ws.set_column('A:A', 15, date_fmt)
        self._add_summary_table(ws, sheet_name, final_df, 'Erlang', 20, 9)
        max_row = len(final_df)

        c1_prim = workbook.add_chart({'type': 'line'})

        if self.selected_region != "South":
            if total_col in final_df.columns:
                c_idx = list(final_df.columns).index(total_col)
                c1_prim.add_series({'name': [sheet_name, 0, c_idx], 'categories': [sheet_name, 1, 0, max_row, 0], 'values': [sheet_name, 1, c_idx, max_row, c_idx]})

            c1_prim.set_title({'name': f'Data and Voice Traffic Behavior of {adj} Cluster – Total Voice Traffic (2G/3G/4G)'})
            c1_prim.set_y_axis({'name': f'Total {adj} Cluster Voice Traffic (Erlang)'})
            c1_prim.set_x_axis({'name': 'Date', 'date_axis': True})
            c1_prim.set_legend({'position': 'bottom'})

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


def build_excel_report(target_base_dir, region, periods=None):
    """periods: optional {"pre": (start, end), "post": (start, end)} of date
    objects; adds a Pre/Post traffic summary table to every sheet."""
    builder = ReportBuilder(target_base_dir, region, periods)
    base_folder = os.path.join(target_base_dir, "Sales Report", region)
    output_path = os.path.join(base_folder, f"{region}_Region_Final_Sales_Report.xlsx")

    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    workbook = writer.book
    builder.workbook = workbook
    date_fmt = workbook.add_format({'num_format': 'm/d/yyyy'})

    if region == "Central":
        builder.process_2g_voice_sheet(base_folder, writer, workbook)
        builder.process_3g_voice_sheet(base_folder, writer, workbook)
        builder.process_3g_data_sheet(base_folder, writer, workbook)
        builder.process_4g_voice_sheet(base_folder, writer, workbook)
        builder.process_4g_data_sheet(base_folder, writer, workbook)

    else:
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
                builder._merge_and_chart_south_north(dfs, writer, workbook, tech, metric_name, unit, date_fmt)

        builder.process_split_volte_sheets(base_folder, writer, workbook)
        builder.process_total_data_sheet(base_folder, writer, workbook)
        builder.process_total_voice_sheet(base_folder, writer, workbook)

    writer.close()
    return output_path
