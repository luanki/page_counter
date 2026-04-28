#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Document Page Counter
A simple desktop app to count pages of Word and PDF files in a folder.
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime

try:
    from PyPDF2 import PdfReader
except ImportError:
    PdfReader = None

try:
    from openpyxl import Workbook
except ImportError:
    Workbook = None


def get_word_page_count(filepath):
    """
    Extract page count from .docx file.
    Word stores document statistics in docProps/app.xml.
    Returns int page count or None if unavailable.
    """
    try:
        with zipfile.ZipFile(filepath, 'r') as z:
            if 'docProps/app.xml' not in z.namelist():
                return None
            with z.open('docProps/app.xml') as app_xml:
                tree = ET.parse(app_xml)
                root = tree.getroot()
                # Namespaces in app.xml
                ns = {
                    'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
                }
                pages_elem = root.find('.//ep:Pages', ns)
                if pages_elem is not None and pages_elem.text:
                    return int(pages_elem.text)
        return None
    except Exception:
        return None


def get_pdf_page_count(filepath):
    """Extract page count from PDF file."""
    if PdfReader is None:
        return None
    try:
        reader = PdfReader(filepath)
        return len(reader.pages)
    except Exception:
        return None


def scan_files(folder_path, read_word=True, read_pdf=True):
    """
    Generator that yields (filename, file_type, page_count, status)
    for each supported file found.
    """
    if not os.path.isdir(folder_path):
        return

    supported_exts = []
    if read_word:
        supported_exts.append('.docx')
    if read_pdf:
        supported_exts.append('.pdf')

    all_files = []
    for root_dir, _, files in os.walk(folder_path):
        for f in files:
            ext = os.path.splitext(f)[1].lower()
            if ext in supported_exts:
                all_files.append((root_dir, f, ext))

    total = len(all_files)
    for idx, (root_dir, filename, ext) in enumerate(all_files, start=1):
        filepath = os.path.join(root_dir, filename)
        rel_path = os.path.relpath(filepath, folder_path)

        if ext == '.docx':
            pages = get_word_page_count(filepath)
            file_type = 'Word'
        elif ext == '.pdf':
            pages = get_pdf_page_count(filepath)
            file_type = 'PDF'
        else:
            pages = None
            file_type = 'Unknown'

        if pages is None:
            status = '无法读取'
        else:
            status = f'{pages} 页'

        yield {
            'index': idx,
            'total': total,
            'filename': filename,
            'rel_path': rel_path,
            'file_type': file_type,
            'pages': pages,
            'status': status,
            'filepath': filepath
        }


class PageCounterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("文档页数统计工具")
        self.root.geometry("900x650")
        self.root.minsize(800, 500)

        self.results = []
        self.scanning = False

        self._build_ui()

    def _build_ui(self):
        # Top frame: folder selection
        top_frame = ttk.Frame(self.root, padding=10)
        top_frame.pack(fill=tk.X)

        ttk.Label(top_frame, text="选择文件夹:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.folder_var = tk.StringVar()
        ttk.Entry(top_frame, textvariable=self.folder_var, width=60).grid(row=0, column=1, padx=5)
        ttk.Button(top_frame, text="浏览...", command=self._browse_folder).grid(row=0, column=2, padx=5)

        # Options frame
        opts_frame = ttk.LabelFrame(self.root, text="文件类型选项", padding=10)
        opts_frame.pack(fill=tk.X, padx=10, pady=5)

        self.read_word_var = tk.BooleanVar(value=True)
        self.read_pdf_var = tk.BooleanVar(value=True)

        ttk.Checkbutton(opts_frame, text="Word 文档 (.docx)", variable=self.read_word_var).pack(side=tk.LEFT, padx=10)
        ttk.Checkbutton(opts_frame, text="PDF 文件 (.pdf)", variable=self.read_pdf_var).pack(side=tk.LEFT, padx=10)

        # Action buttons
        btn_frame = ttk.Frame(self.root, padding=10)
        btn_frame.pack(fill=tk.X)

        self.start_btn = ttk.Button(btn_frame, text="开始扫描", command=self._start_scan)
        self.start_btn.pack(side=tk.LEFT, padx=5)

        self.export_btn = ttk.Button(btn_frame, text="导出 Excel", command=self._export_excel, state=tk.DISABLED)
        self.export_btn.pack(side=tk.LEFT, padx=5)

        self.clear_btn = ttk.Button(btn_frame, text="清空结果", command=self._clear_results)
        self.clear_btn.pack(side=tk.LEFT, padx=5)

        # Progress bar
        prog_frame = ttk.Frame(self.root, padding=(10, 0))
        prog_frame.pack(fill=tk.X)

        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(prog_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=5)

        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(prog_frame, textvariable=self.status_var).pack(anchor=tk.W)

        # Results treeview
        tree_frame = ttk.Frame(self.root, padding=10)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("序号", "文件名", "类型", "页数", "状态")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")

        self.tree.heading("序号", text="序号")
        self.tree.heading("文件名", text="文件名")
        self.tree.heading("类型", text="类型")
        self.tree.heading("页数", text="页数")
        self.tree.heading("状态", text="状态")

        self.tree.column("序号", width=60, anchor=tk.CENTER)
        self.tree.column("文件名", width=400, anchor=tk.W)
        self.tree.column("类型", width=80, anchor=tk.CENTER)
        self.tree.column("页数", width=80, anchor=tk.CENTER)
        self.tree.column("状态", width=100, anchor=tk.CENTER)

        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # Summary label
        self.summary_var = tk.StringVar(value="")
        ttk.Label(self.root, textvariable=self.summary_var, padding=10).pack(anchor=tk.W)

    def _browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_var.set(folder)

    def _clear_results(self):
        self.results.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.progress_var.set(0)
        self.status_var.set("就绪")
        self.summary_var.set("")
        self.export_btn.config(state=tk.DISABLED)

    def _start_scan(self):
        folder = self.folder_var.get().strip()
        if not folder:
            messagebox.showwarning("提示", "请先选择要扫描的文件夹！")
            return
        if not os.path.isdir(folder):
            messagebox.showerror("错误", "所选路径不存在或不是文件夹！")
            return
        if not self.read_word_var.get() and not self.read_pdf_var.get():
            messagebox.showwarning("提示", "请至少选择一种文件类型！")
            return

        self.results.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.export_btn.config(state=tk.DISABLED)
        self.progress_var.set(0)
        self.status_var.set("扫描中...")
        self.start_btn.config(state=tk.DISABLED)
        self.scanning = True

        thread = threading.Thread(target=self._scan_worker, args=(folder,), daemon=True)
        thread.start()

    def _scan_worker(self, folder):
        read_word = self.read_word_var.get()
        read_pdf = self.read_pdf_var.get()

        try:
            for info in scan_files(folder, read_word, read_pdf):
                if not self.scanning:
                    break
                self.results.append(info)
                pct = (info['index'] / info['total'] * 100) if info['total'] > 0 else 100
                self.root.after(0, self._update_ui, info, pct)
        except Exception as e:
            self.root.after(0, lambda: self.status_var.set(f"出错: {e}"))
        finally:
            self.root.after(0, self._scan_done)

    def _update_ui(self, info, pct):
        self.tree.insert("", tk.END, values=(
            info['index'],
            info['rel_path'],
            info['file_type'],
            info['pages'] if info['pages'] is not None else '-',
            info['status']
        ))
        self.progress_var.set(pct)
        self.status_var.set(f"正在扫描: {info['filename']} ({info['index']}/{info['total']})")
        self.tree.yview_moveto(1.0)

    def _scan_done(self):
        self.scanning = False
        self.start_btn.config(state=tk.NORMAL)
        total_files = len(self.results)
        total_pages = sum(r['pages'] for r in self.results if r['pages'] is not None)
        failed = sum(1 for r in self.results if r['pages'] is None)
        self.status_var.set(f"扫描完成: 共 {total_files} 个文件")
        self.summary_var.set(f"总计: {total_files} 个文件 | 成功读取页数: {total_files - failed} 个 | 总页数: {total_pages} 页")
        if self.results:
            self.export_btn.config(state=tk.NORMAL)

    def _export_excel(self):
        if not self.results:
            messagebox.showwarning("提示", "没有可导出的数据！")
            return
        if Workbook is None:
            messagebox.showerror("错误", "缺少 openpyxl 库，无法导出 Excel！")
            return

        default_name = f"文档页数统计_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile=default_name
        )
        if not filepath:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "统计结果"
            ws.append(["序号", "文件名", "文件类型", "页数", "状态", "完整路径"])
            for idx, r in enumerate(self.results, start=1):
                ws.append([
                    idx,
                    r['filename'],
                    r['file_type'],
                    r['pages'] if r['pages'] is not None else '-',
                    r['status'],
                    r['filepath']
                ])
            # Adjust column widths
            ws.column_dimensions['A'].width = 10
            ws.column_dimensions['B'].width = 50
            ws.column_dimensions['C'].width = 15
            ws.column_dimensions['D'].width = 10
            ws.column_dimensions['E'].width = 15
            ws.column_dimensions['F'].width = 80
            wb.save(filepath)
            messagebox.showinfo("成功", f"Excel 已导出到:\n{filepath}")
        except Exception as e:
            messagebox.showerror("导出失败", f"导出 Excel 时出错:\n{e}")


def main():
    root = tk.Tk()
    app = PageCounterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
