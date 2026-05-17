"""Tkinter GUI 前端。"""

import json
import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

from .core import DEFAULT_PLACEHOLDER_STYLES, generate_documents, parse_map_args


class Excel2WordGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel 2 Word 模板填充")
        self.geometry("760x640")
        self.resizable(True, True)

        self.excel_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.group_by = tk.StringVar()
        self.output_name = tk.StringVar(value="{ROW_INDEX}.docx")
        self.repeat_marker = tk.StringVar(value="#repeat")

        self.create_widgets()

    def create_widgets(self):
        frame = tk.Frame(self, padx=12, pady=12)
        frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(frame, text="Excel 文件 (.xlsx)").grid(row=0, column=0, sticky="w")
        tk.Entry(frame, textvariable=self.excel_path, width=72).grid(row=1, column=0, sticky="we", columnspan=3)
        tk.Button(frame, text="选择", command=self.select_excel).grid(row=1, column=3, padx=8)

        tk.Label(frame, text="Word 模板 (.docx)").grid(row=2, column=0, sticky="w", pady=(12, 0))
        tk.Entry(frame, textvariable=self.template_path, width=72).grid(row=3, column=0, sticky="we", columnspan=3)
        tk.Button(frame, text="选择", command=self.select_template).grid(row=3, column=3, padx=8)

        tk.Label(frame, text="输出目录").grid(row=4, column=0, sticky="w", pady=(12, 0))
        tk.Entry(frame, textvariable=self.output_dir, width=72).grid(row=5, column=0, sticky="we", columnspan=3)
        tk.Button(frame, text="选择", command=self.select_output).grid(row=5, column=3, padx=8)

        tk.Label(frame, text="Excel 列映射（每行一个：表头=占位符）").grid(row=6, column=0, sticky="w", pady=(12, 0), columnspan=4)
        self.mapping_text = scrolledtext.ScrolledText(frame, height=8, width=94)
        self.mapping_text.insert(tk.END, "盒号=BOX_NO\n文号=DOC_NO\n责任者=OWNER\n文件题名=TITLE\n")
        self.mapping_text.grid(row=7, column=0, columnspan=4, sticky="nsew")

        tk.Label(frame, text="按此列分组生成一个 Word 文档（可选）").grid(row=8, column=0, sticky="w", pady=(12, 0), columnspan=4)
        tk.Entry(frame, textvariable=self.group_by, width=32).grid(row=9, column=0, columnspan=2, sticky="w")
        tk.Label(frame, text="输出文件名模板").grid(row=8, column=2, sticky="w", pady=(12, 0))
        tk.Entry(frame, textvariable=self.output_name, width=32).grid(row=9, column=2, columnspan=2, sticky="w")

        tk.Label(frame, text="表格复制模板标记（例如 {{#repeat}}）").grid(row=10, column=0, sticky="w", pady=(12, 0), columnspan=2)
        tk.Entry(frame, textvariable=self.repeat_marker, width=20).grid(row=11, column=0, sticky="w")

        button_frame = tk.Frame(frame)
        button_frame.grid(row=12, column=0, columnspan=4, pady=18)
        tk.Button(button_frame, text="生成文档", command=self.on_generate, width=14, bg="#4a8cff", fg="white").pack(side=tk.LEFT, padx=6)
        tk.Button(button_frame, text="清空日志", command=self.clear_log, width=12).pack(side=tk.LEFT, padx=6)

        tk.Label(frame, text="日志输出").grid(row=13, column=0, sticky="w")
        self.log_area = scrolledtext.ScrolledText(frame, height=12, width=94, state=tk.DISABLED)
        self.log_area.grid(row=14, column=0, columnspan=4, sticky="nsew")

        frame.grid_rowconfigure(7, weight=1)
        frame.grid_rowconfigure(14, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(1, weight=1)
        frame.grid_columnconfigure(2, weight=1)
        frame.grid_columnconfigure(3, weight=0)

    def select_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx")])
        if path:
            self.excel_path.set(path)

    def select_template(self):
        path = filedialog.askopenfilename(filetypes=[("Word 文件", "*.docx")])
        if path:
            self.template_path.set(path)

    def select_output(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir.set(path)

    def log(self, message: str) -> None:
        self.log_area.config(state=tk.NORMAL)
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state=tk.DISABLED)

    def clear_log(self) -> None:
        self.log_area.config(state=tk.NORMAL)
        self.log_area.delete("1.0", tk.END)
        self.log_area.config(state=tk.DISABLED)

    def read_mapping(self) -> List[str]:
        text = self.mapping_text.get("1.0", tk.END).strip()
        return [line.strip() for line in text.splitlines() if line.strip()]

    def on_generate(self):
        excel_path = self.excel_path.get().strip()
        template_path = self.template_path.get().strip()
        output_dir = self.output_dir.get().strip()
        mappings = self.read_mapping()

        if not excel_path or not template_path or not output_dir:
            messagebox.showwarning("参数缺失", "请先选择 Excel、Word 模板和输出目录。")
            return

        try:
            mapping_dict = parse_map_args(mappings)
            output_pattern = self.output_name.get().strip() or "{ROW_INDEX}.docx"
            group_by = self.group_by.get().strip() or None
            files = generate_documents(
                excel_path=excel_path,
                template_path=template_path,
                output_dir=output_dir,
                mapping=mapping_dict,
                output_name_pattern=output_pattern,
                group_by=group_by,
                repeat_marker=self.repeat_marker.get().strip() or "#repeat",
                placeholder_styles=DEFAULT_PLACEHOLDER_STYLES,
            )
            self.log(f"成功生成 {len(files)} 个文档：")
            for path in files:
                self.log(f"  - {path}")
            messagebox.showinfo("完成", f"已生成 {len(files)} 个文档。")
        except Exception as exc:
            self.log(f"错误: {exc}")
            messagebox.showerror("生成失败", str(exc))


def main() -> None:
    app = Excel2WordGUI()
    app.mainloop()
