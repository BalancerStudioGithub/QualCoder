import sys
import ctypes
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import pandas as pd
from datetime import datetime
import os
import re

if sys.platform.startswith('win'):
    kernel32 = ctypes.windll.kernel32
    user32 = ctypes.windll.user32
    hWnd = kernel32.GetConsoleWindow()
    if hWnd:
        user32.ShowWindow(hWnd, 0)


def convert_timeline_format(timeline_str, default_end_date):
    try:
        if "~" in timeline_str:
            start_date_str, end_date_str = timeline_str.split("~")
            start_date_str = start_date_str.strip()
            end_date_str = end_date_str.strip() or default_end_date.strip()
            try:
                start_date = datetime.strptime(start_date_str, "%Y/%m/%d")
                start_date_formatted = start_date.strftime('%Y-%m-%d')
            except ValueError:
                start_date = datetime.strptime(start_date_str, "%Y/%m")
                start_date_formatted = start_date.strftime('%Y-%m-01')
            try:
                end_date = datetime.strptime(end_date_str, "%Y/%m/%d")
                end_date_formatted = end_date.strftime('%Y-%m-%d')
            except ValueError:
                end_date = datetime.strptime(end_date_str, "%Y/%m")
                end_date_formatted = end_date.strftime('%Y-%m-01')

            return start_date_formatted, end_date_formatted
        else:
            start_date_str = timeline_str.strip()
            try:
                start_date = datetime.strptime(start_date_str, "%Y/%m/%d")
                start_date_formatted = start_date.strftime('%Y-%m-%d')
            except ValueError:
                start_date = datetime.strptime(start_date_str, "%Y/%m")
                start_date_formatted = start_date.strftime('%Y-%m-01')

            return start_date_formatted, default_end_date
    except ValueError:
        return "Invalid Format", "Invalid Format"


def categorize_text(text, categories):
    for category in categories:
        if category in text:
            return category
    return text


class DataProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Processor and Mermaid Gantt Chart Generator")
        self.root.geometry("1000x600")
        self.data = None
        self.default_end_date = tk.StringVar(value="2025/01")
        self.category_keywords = tk.StringVar(value="Software|Hardware|Theory|Traditional Skills|Medium|Others")
        self.milestone_keywords = tk.StringVar(value="Theory|other")
        self.mermaid_theme_option = tk.BooleanVar(value=False)
        self.count_thresholds = {'crit': tk.StringVar(value="5"), 'active': tk.StringVar(value="3"),
                                 'done': tk.StringVar(value="2")}
        self.year_month_switch = tk.BooleanVar(value=True)

        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Top row frame
        top_row_frame = ttk.Frame(main_frame)
        top_row_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

        tk.Button(top_row_frame, text="Import CSV", command=self.import_csv).pack(side=tk.LEFT, padx=5)
        tk.Button(top_row_frame, text="Export to XLSX", command=self.export_to_xlsx).pack(side=tk.LEFT, padx=5)
        tk.Checkbutton(top_row_frame, text="Enable Custom Theme", variable=self.mermaid_theme_option).pack(side=tk.LEFT,
                                                                                                           padx=5)
        tk.Label(top_row_frame, text="Default End Date (YYYY/M):").pack(side=tk.LEFT, padx=5)
        tk.Entry(top_row_frame, textvariable=self.default_end_date, width=8).pack(side=tk.LEFT, padx=5)
        tk.Label(top_row_frame, text="Thresholds for Importance Levels:").pack(side=tk.LEFT, padx=5)

        for label, var in self.count_thresholds.items():
            tk.Label(top_row_frame, text=label).pack(side=tk.LEFT, padx=5)
            tk.Entry(top_row_frame, textvariable=var, width=3).pack(side=tk.LEFT, padx=5)

        tk.Checkbutton(top_row_frame, text="Include YearMonth", variable=self.year_month_switch).pack(side=tk.LEFT,
                                                                                                      padx=5)

        # Second row frame
        second_row_frame = ttk.Frame(main_frame)
        second_row_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

        tk.Label(second_row_frame, text="Category Keywords (?|?):").pack(side=tk.LEFT, padx=5)
        tk.Entry(second_row_frame, textvariable=self.category_keywords, width=75).pack(side=tk.LEFT, padx=5)
        tk.Label(second_row_frame, text="Milestone Keywords (?|?):").pack(side=tk.LEFT, padx=5)
        tk.Entry(second_row_frame, textvariable=self.milestone_keywords, width=50).pack(side=tk.LEFT, padx=5)

        # Left frame
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.preview_tree = ttk.Treeview(left_frame)
        self.preview_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=self.preview_tree.yview)
        self.scrollbar.pack(side=tk.LEFT, fill=tk.Y)

        self.preview_tree.configure(yscrollcommand=self.scrollbar.set)

        # Right frame
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.generate_button = tk.Button(right_frame, text="Generate Mermaid Gantt",
                                         command=self.generate_mermaid_gantt, state='disabled')
        self.generate_button.pack(side=tk.TOP, pady=10)

        self.mermaid_text = scrolledtext.ScrolledText(right_frame, wrap=tk.WORD, width=80, height=20)
        self.mermaid_text.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.root.bind("<Configure>", self.on_maximize)

    def import_csv(self):
        filepath = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not filepath:
            return
        try:
            self.data = pd.read_csv(filepath, sep='\t')
            self.convert_timeline_column()
            self.categorize_column()
            self.add_importance_column()
            self.parse_code_memo_column()
            self.preview_data(self.data)
            self.generate_button['state'] = 'normal'
            self.update_layout()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import file: {e}")

    def convert_timeline_column(self):
        if 'Timeline' in self.data.columns:
            self.data[['StartTime', 'EndTime']] = self.data['Timeline'].apply(
                lambda x: pd.Series(convert_timeline_format(x, self.default_end_date.get())))

    def categorize_column(self):
        categories = [cat.strip() for cat in self.category_keywords.get().split("|")]
        patterns = {cat: re.compile(r'\b' + re.escape(cat) + r'\b') for cat in categories}

        if 'Category' in self.data.columns:
            self.data['Category'] = self.data['Category'].apply(lambda x:
                                                                next((category for category in categories if
                                                                      patterns[category].search(x)), x))

    def add_importance_column(self):
        if 'Count' in self.data.columns:
            thresholds = {key: int(var.get()) for key, var in self.count_thresholds.items()}
            milestone_keywords = self.milestone_keywords.get().split("|")
            self.data['Importance'] = self.data.apply(lambda row:
                                                      'milestone' if any(keyword in row['Category'] for keyword in
                                                                         milestone_keywords) else
                                                      'crit' if row['Count'] >= thresholds['crit'] else
                                                      'active' if row['Count'] >= thresholds['active'] else
                                                      'done' if row['Count'] >= thresholds['done'] else '', axis=1)

    def parse_code_memo_column(self):
        if 'Code Memo' in self.data.columns:
            self.data['Link'] = self.data['Code Memo'].apply(self.extract_link_from_code_memo)

    def extract_link_from_code_memo(self, code_memo):
        url_pattern = r'https?://[^\s]+'
        match = re.search(url_pattern, code_memo)
        if match:
            return match.group(0)
        return None

    def preview_data(self, df):
        categories = [cat.strip() for cat in self.category_keywords.get().split("|")]
        pattern = '|'.join(f'\\b{re.escape(cat)}\\b' for cat in categories)

        filtered_df = df[df['Category'].str.contains(pattern, regex=True)]

        for i in self.preview_tree.get_children():
            self.preview_tree.delete(i)
        self.preview_tree["columns"] = list(filtered_df.columns)
        fixed_width = 100
        for col in filtered_df.columns:
            self.preview_tree.heading(col, text=col)
            self.preview_tree.column(col, anchor="w")
            if col in ['Timeline', 'StartTime', 'EndTime']:
                self.preview_tree.column(col, width=fixed_width)
            else:
                max_width = max([len(str(item)) * 7 for item in filtered_df[col]])
                self.preview_tree.column(col, width=max_width - 95)
        self.preview_tree.column('#0', width=0, stretch=tk.NO)
        for col in filtered_df.columns:
            self.preview_tree.column(col, anchor="w", minwidth=fixed_width)

        for index, row in filtered_df.iterrows():
            importance = f"{row['Importance']}, " if row['Importance'] else ""
            self.preview_tree.insert("", "end", text="", values=(row['Timeline'],) + tuple(row[1:]))

    def generate_mermaid_gantt(self):
        if self.data is None:
            messagebox.showwarning("Warning", "No data to process. Please import a CSV file first.")
            return
        filtered_data = self.data[self.data['Category'].isin(self.category_keywords.get().split("|"))]
        mermaid_code = self.generate_mermaid_gantt_code(filtered_data)
        self.mermaid_text.delete('1.0', tk.END)
        self.mermaid_text.insert(tk.END, mermaid_code)

    def generate_mermaid_gantt_code(self, df):
        sections = df['Category'].unique()
        mermaid_code = "gantt\n"
        mermaid_code += "title Media Development Timeline\ndateFormat YYYY-MM-DD\naxisFormat %Y\n"
        if self.mermaid_theme_option.get():
            mermaid_code += "%%{init: { 'theme': 'base', 'themeVariables': {\n"
            mermaid_code += " 'primaryColor': '#D3D3D3', 'primaryTextColor': '#000000', 'primaryBorderColor': '#A9A9A1',\n"
            mermaid_code += " 'lineColor': '#eee', 'secondaryColor': '#f0e68c', 'tertiaryColor': '#C0C0C0'\n"
            mermaid_code += "}, 'gantt': {\n"
            mermaid_code += " 'topAxis': 1,\n"
            mermaid_code += " 'titleTopMargin': 25,\n"
            mermaid_code += " 'barHeight': 20,\n"
            mermaid_code += " 'barGap': 4,\n"
            mermaid_code += " 'topPadding': 75,\n"
            mermaid_code += " 'rightPadding': 5,\n"
            mermaid_code += " 'leftPadding': 5,\n"
            mermaid_code += " 'gridLineStartPadding': 100,\n"
            mermaid_code += " 'fontSize': 20,\n"
            mermaid_code += " 'sectionFontSize': 25,\n"
            mermaid_code += " 'numberSectionStyles': 2\n"
            mermaid_code += "} } }%%\n"
        mermaid_code += "todayMarker stroke-width:1px,stroke:#d8d8d8,opacity:0.8\n"
        for section in sections:
            mermaid_code += f"section {section}\n"
            for index, row in df[df['Category'] == section].iterrows():
                importance = row['Importance']
                start_date, end_date = row['StartTime'], row['EndTime']
                task_name = row['Code Name']
                link = row['Link']
                yearmonth = row['Timeline']

                if self.year_month_switch.get():
                    task_name += f"_{yearmonth}"

                if importance == 'milestone':
                    mermaid_code += f" {task_name} : {importance}, t{index}, {start_date} , {start_date}\n"
                    if link:
                        mermaid_code += f"click t{index} href \"{link}\"\n"
                elif importance:
                    mermaid_code += f" {task_name} : {importance}, t{index}, {start_date} , {end_date}\n"
                    if link:
                        mermaid_code += f"click t{index} href \"{link}\"\n"
                else:
                    mermaid_code += f" {task_name} : t{index}, {start_date} , {end_date}\n"
                    if link:
                        mermaid_code += f"click t{index} href \"{link}\"\n"
        return mermaid_code

    def update_layout(self):
        total_width = self.root.winfo_width()
        left_width = int(total_width * 0.7)
        right_width = total_width - left_width
        self.preview_tree.master.config(width=left_width)
        self.mermaid_text.master.config(width=right_width)

    def on_maximize(self, event):
        if self.root.state() == 'zoomed':
            self.update_layout()

    def export_to_xlsx(self):
        if self.data is None:
            messagebox.showwarning("Warning", "No data to process. Please import a CSV file first.")
            return
        filtered_data = self.data[self.data['Category'].isin(self.category_keywords.get().split("|"))]
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if not filepath:
            return
        try:
            filtered_data.to_excel(filepath, index=False)
            messagebox.showinfo("Success", f"Data exported to {os.path.basename(filepath)} successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export data: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = DataProcessor(root)
    root.mainloop()