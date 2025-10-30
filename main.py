import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import pandas as pd
import xlrd  # 保证已在requirements.txt
import warnings


class ExcelProcessorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("joco")
        self.geometry("1000x640")

        # 数据状态
        self.loaded_file_path = None
        self.loaded_sheets = []
        self.current_df = None
        self.preview_df = None

        # UI
        self._build_ui()

    def _build_ui(self):
        # 顶部：文件选择 + 工作表选择
        top_frame = ttk.Frame(self, padding=8)
        top_frame.pack(fill=tk.X)

        self.file_label_var = tk.StringVar(value="未选择文件")
        ttk.Button(top_frame, text="选择Excel文件", command=self.on_select_file).pack(side=tk.LEFT)
        ttk.Label(top_frame, textvariable=self.file_label_var).pack(side=tk.LEFT, padx=8)

        ttk.Label(top_frame, text="工作表:").pack(side=tk.LEFT, padx=(16, 4))
        self.sheet_combo = ttk.Combobox(top_frame, state="disabled", width=24)
        self.sheet_combo.pack(side=tk.LEFT)
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_change)

        # 中部：列选择与增量设置
        middle_frame = ttk.Frame(self, padding=8)
        middle_frame.pack(fill=tk.X)

        # 列选择（多选），带映射
        self.template_headers = []   # 存储模板字段列表
        self.col_mapping_vars = {}   # key:源列, value:tk.StringVar
        self.mapping_widgets = {}    # 动态Combobox控件，便于后续清理

        # 读取模板表头以供选项
        try:
            template_path = os.path.join(os.path.dirname(__file__), "CPS团链产业777.xls")
            workbook = xlrd.open_workbook(template_path)
            sheet = workbook.sheet_by_index(0)
            self.template_headers = sheet.row_values(0)
        except Exception:
            self.template_headers = []

        left_box = ttk.Labelframe(middle_frame, text="选择要提取的列", padding=8)
        left_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.mapping_canvas = tk.Canvas(left_box)
        self.mapping_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(left_box, orient="vertical", command=self.mapping_canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.mapping_canvas.configure(yscrollcommand=scrollbar.set)
        self.mapping_inner = ttk.Frame(self.mapping_canvas)
        self.mapping_canvas.create_window((0, 0), window=self.mapping_inner, anchor='nw')
        self.mapping_inner.bind("<Configure>", lambda e: self.mapping_canvas.configure(scrollregion=self.mapping_canvas.bbox("all")))
        # 注意：self.columns_listbox 暂不再直接展示，推荐用 mapping_inner 实现UI

        # 操作区
        right_box = ttk.Labelframe(middle_frame, text="列增量设置", padding=8)
        right_box.pack(side=tk.LEFT, fill=tk.Y, padx=(8, 0))

        ttk.Label(right_box, text="选择要+0.001的列").grid(row=0, column=0, sticky=tk.W)
        self.increment_col_combo = ttk.Combobox(right_box, state="disabled", width=24)
        self.increment_col_combo.grid(row=1, column=0, sticky=tk.W)

        ttk.Label(right_box, text="增量(默认0.001)").grid(row=2, column=0, sticky=tk.W, pady=(8, 0))
        self.increment_value_var = tk.StringVar(value="0.001")
        self.increment_entry = ttk.Entry(right_box, textvariable=self.increment_value_var, width=12)
        self.increment_entry.grid(row=3, column=0, sticky=tk.W)

        ttk.Button(right_box, text="应用提取与增量", command=self.on_apply).grid(row=4, column=0, sticky=tk.W, pady=(12, 0))

        for i in range(5):
            right_box.grid_rowconfigure(i, pad=2)

        # 预览表格
        preview_box = ttk.Labelframe(self, text="预览(前200行)", padding=8)
        preview_box.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        self.tree = ttk.Treeview(preview_box, columns=(), show="headings")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        vsb = ttk.Scrollbar(preview_box, orient="vertical", command=self.tree.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb = ttk.Scrollbar(preview_box, orient="horizontal", command=self.tree.xview)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # 底部：导出
        bottom_frame = ttk.Frame(self, padding=8)
        bottom_frame.pack(fill=tk.X)
        ttk.Button(bottom_frame, text="导出为Excel...", command=self.on_export).pack(side=tk.RIGHT)

    # 事件处理
    def on_select_file(self):
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")],
        )
        if not file_path:
            return

        try:
            xls = pd.ExcelFile(file_path)
            self.loaded_file_path = file_path
            self.loaded_sheets = xls.sheet_names
            self.file_label_var.set(os.path.basename(file_path))

            self.sheet_combo.configure(state="readonly", values=self.loaded_sheets)
            if self.loaded_sheets:
                self.sheet_combo.set(self.loaded_sheets[0])
                self._load_sheet(self.loaded_sheets[0])
        except Exception as e:
            messagebox.showerror("读取失败", f"无法读取文件:\n{e}")

    def on_sheet_change(self, event=None):
        sheet = self.sheet_combo.get()
        if sheet:
            self._load_sheet(sheet)

    def _load_sheet(self, sheet_name: str):
        if not self.loaded_file_path:
            return
        try:
            df = pd.read_excel(self.loaded_file_path, sheet_name=sheet_name)
            self.current_df = df
            self._refresh_columns(df)
            self._set_preview(df.head(200))
        except Exception as e:
            messagebox.showerror("读取失败", f"无法读取工作表:\n{e}")

    def _refresh_columns(self, df: pd.DataFrame):
        cols = list(map(str, df.columns))
        # 清理UI
        for widget in self.mapping_inner.winfo_children():
            widget.destroy()
        self.col_mapping_vars = {}  # 重建变量
        self.mapping_widgets = {}
        self.col_checkbox_vars = {}

        # 每列渲染一行
        for i, c in enumerate(cols):
            var = tk.BooleanVar(value=False)
            self.col_checkbox_vars[c] = var
            chk = ttk.Checkbutton(self.mapping_inner, text=c, variable=var)
            chk.grid(row=i, column=0, sticky=tk.W)

            lbl = ttk.Label(self.mapping_inner, text="→", width=2)
            lbl.grid(row=i, column=1)

            cmb_var = tk.StringVar()
            cmb = ttk.Combobox(self.mapping_inner, textvariable=cmb_var, state="readonly", width=28)
            cmb["values"] = self.template_headers
            if self.template_headers:
                cmb_var.set("")  # 默认空
            cmb.grid(row=i, column=2, sticky=tk.W)

            self.col_mapping_vars[c] = cmb_var
            self.mapping_widgets[c] = cmb

        self.increment_col_combo.configure(state="readonly", values=cols)
        if cols:
            self.increment_col_combo.set(cols[0])

    def _set_preview(self, df: pd.DataFrame):
        self.preview_df = df
        # 清空 Treeview
        for col in self.tree["columns"]:
            self.tree.heading(col, text="")
            self.tree.column(col, width=100)
        self.tree.delete(*self.tree.get_children())

        if df is None or df.empty:
            self.tree.configure(columns=())
            return

        cols = list(map(str, df.columns))
        self.tree.configure(columns=cols)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=max(80, min(240, int(8 * (len(str(c)) + 3)))))

        # 插入行（最多200行）
        for _, row in df.iterrows():
            values = [row.get(c) for c in df.columns]
            self.tree.insert("", tk.END, values=values)

    def on_apply(self):
        if self.current_df is None:
            messagebox.showwarning("提示", "请先选择并加载Excel文件")
            return

        # 获取勾选列和映射
        selected_cols = [col for col, var in self.col_checkbox_vars.items() if var.get()]
        if not selected_cols:
            messagebox.showwarning("提示", "请至少选择一列进行提取")
            return

        mapping = {}
        used_targets = set()
        for col in selected_cols:
            target = self.col_mapping_vars.get(col).get().strip()
            if target:
                if target in used_targets:
                    messagebox.showerror("错误", f"目标列{target}已被多次映射，请调整后再继续！")
                    return
                mapping[col] = target
                used_targets.add(target)
            else:
                mapping[col] = None

        df = self.current_df[selected_cols].copy()

        # 增量逻辑不变
        target_col = self.increment_col_combo.get()
        inc_str = self.increment_value_var.get().strip()
        try:
            inc_value = float(inc_str)
        except Exception:
            messagebox.showerror("错误", f"增量值无效: {inc_str}")
            return

        if target_col:
            if target_col not in df.columns:
                if target_col in self.current_df.columns:
                    df[target_col] = self.current_df[target_col]
                else:
                    messagebox.showerror("错误", f"未找到列: {target_col}")
                    return
            def add_inc(v):
                try:
                    return float(v) + inc_value
                except Exception:
                    return v
            df[target_col] = df[target_col].map(add_inc)

        # 预览阶段先不调整顺序，仅显示数据
        self._set_preview(df.head(200))
        self.preview_df = df
        self._active_col_mapping = mapping  # 供导出用

    def on_export(self):
        if self.preview_df is None or self.preview_df.empty:
            messagebox.showwarning("提示", "没有可导出的数据，请先应用提取与增量")
            return
        file_path = filedialog.asksaveasfilename(
            title="导出为Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not file_path:
            return
        try:
            # 以 xls 模板获取表头
            template_path = os.path.join(os.path.dirname(__file__), "CPS团链产业777.xls")
            workbook = xlrd.open_workbook(template_path)
            sheet = workbook.sheet_by_index(0)
            template_headers = sheet.row_values(0)
        except Exception:
            template_headers = []

        mapping = getattr(self, '_active_col_mapping', {})
        data_df = self.preview_df.copy()

        # 1. 生成映射后DataFrame，模板表头优先顺序
        if mapping and template_headers:
            result = pd.DataFrame(columns=template_headers)
            unmapped_cols = []
            for src_col, tgt_col in mapping.items():
                if tgt_col and tgt_col in template_headers:
                    # 填写到目标template字段
                    result[tgt_col] = data_df[src_col]
                else:
                    unmapped_cols.append(src_col)
            # 未在模板字段的源数据，原名放末尾
            for c in unmapped_cols:
                result[c] = data_df[c]
            result.to_excel(file_path, index=False)
        else:
            # 回退原始导出
            data_df.to_excel(file_path, index=False)

        messagebox.showinfo("成功", f"已导出到:\n{file_path}")


def main():
    app = ExcelProcessorApp()
    app.mainloop()


if __name__ == "__main__":
    main()


