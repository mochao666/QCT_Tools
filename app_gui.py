# -*- coding: utf-8 -*-
"""
QCT 小工具 GUI：导入 PDT / 导入 QCT → 选择 Event（整份 QCT 共用）→ 导出 QCT / 导出 Comments。
导出 QCT 时，两个 Sheet 的第一列自动填充为当前选择的 Event。
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from pdt_reader import read_and_clean_pdt
from qct_data import (
    build_qct_rows_from_pdt,
    read_qct_workbook,
    write_qct_workbook,
    write_comments_workbook,
    EVENT_OPTIONS_LIST,
    EDITABLE_COL_QC_DESC,
)


class QCTToolApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("QCT小工具")
        self.root.minsize(520, 220)
        self.root.geometry("620x260")

        self.sdtm_rows = []
        self.adam_tfl_rows = []
        self._pdt_path = None
        self._qct_path = None  # 导入的 QCT 路径，用于导出时默认文件名

        self._build_ui()

    def _build_ui(self):
        # 导入/导出按钮颜色区分：导入=蓝色，导出=绿色
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Import.TButton", background="#bbdefb")
        style.map("Import.TButton", background=[("active", "#90caf9")])
        style.configure("Export.TButton", background="#c8e6c9")
        style.map("Export.TButton", background=[("active", "#a5d6a7")])

        top = ttk.Frame(self.root, padding=6)
        top.pack(fill=tk.X)
        btn_row = ttk.Frame(top)
        btn_row.pack(fill=tk.X)
        ttk.Button(btn_row, text="导入 PDT", command=self._import_pdt, style="Import.TButton").pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_row, text="导入 QCT", command=self._import_qct, style="Import.TButton").pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_row, text="导出 QCT", command=self._export_qct, style="Export.TButton").pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_row, text="导出Comments", command=self._export_comments, style="Export.TButton").pack(side=tk.LEFT, padx=4)
        self._status = ttk.Label(top, text="请先导入 PDT 或 QCT 文件", foreground="gray")
        self._status.pack(fill=tk.X, pady=(6, 0), anchor=tk.W)

        # Event：整份 QCT 共用，导出时自动填充到两 Sheet 第一列
        edit_frame = ttk.LabelFrame(self.root, text="Event（导出 QCT 时将填充到两个 Sheet 的第一列）", padding=8)
        edit_frame.pack(fill=tk.X, padx=6, pady=8)
        ttk.Label(edit_frame, text="Event:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self._event_var = tk.StringVar(value="")
        self._combo_event = ttk.Combobox(
            edit_frame, textvariable=self._event_var, width=18,
            values=EVENT_OPTIONS_LIST, state="readonly",
        )
        self._combo_event.grid(row=1, column=0, sticky=tk.W, pady=2)

    def _import_pdt(self):
        path = filedialog.askopenfilename(
            title="选择 PDT 文件",
            filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")],
        )
        if not path:
            return
        try:
            pdt_df = read_and_clean_pdt(path)
            event_val = self._event_var.get().strip()
            self.sdtm_rows, self.adam_tfl_rows = build_qct_rows_from_pdt(pdt_df, event_value=event_val)
            self._pdt_path = path
            self._qct_path = None
            self._status.config(
                text=f"已导入\n{os.path.basename(path)}  |  SDTM 行: {len(self.sdtm_rows)}  ADaM/TFL 行: {len(self.adam_tfl_rows)}",
                foreground="black",
            )
            messagebox.showinfo("导入成功", f"PDT 导入成功。\nSDTM 行: {len(self.sdtm_rows)}\nADaM/TFL 行: {len(self.adam_tfl_rows)}")
        except Exception as e:
            messagebox.showerror("导入失败", str(e))

    def _import_qct(self):
        path = filedialog.askopenfilename(
            title="选择 QCT 文件",
            filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")],
        )
        if not path:
            return
        try:
            self.sdtm_rows, self.adam_tfl_rows, _ = read_qct_workbook(path)
            self._qct_path = path
            # 导入 QCT 时不改 Event 选择，仅导入 PDT 时需要选择 Event
            self._status.config(
                text=f"已导入 QCT\n{os.path.basename(path)}  |  SDTM 行: {len(self.sdtm_rows)}  ADaM/TFL 行: {len(self.adam_tfl_rows)}（已过滤 QC results description 为空的行）",
                foreground="black",
            )
        except Exception as e:
            messagebox.showerror("导入 QCT 失败", str(e))

    def _default_export_name(self, suffix):
        """导出默认文件名：有 QCT 路径时用 QCT 文件名（仅 QCT 导出），否则用 PDT 路径推导。"""
        if suffix == "QCT" and self._qct_path:
            return os.path.basename(self._qct_path)
        if self._pdt_path:
            base = os.path.splitext(os.path.basename(self._pdt_path))[0]
            if base.upper().endswith("_PDT"):
                prefix = base[:-4] + "_"
            else:
                prefix = base + "_"
            return prefix + suffix + ".xlsx"
        return None

    def _default_export_dir(self, prefer_qct=False):
        """导出默认目录：优先 QCT 路径（若存在且 prefer_qct），否则 PDT 路径。"""
        if prefer_qct and self._qct_path:
            return os.path.dirname(self._qct_path)
        if self._pdt_path:
            return os.path.dirname(self._pdt_path)
        return None

    def _ask_export_qct_mode(self):
        """导出 QCT 时选择：初版QCT、新增Event 或 终版QCT。返回 'initial'、'append' 或 'final'。"""
        choice = [None]

        def on_initial():
            choice[0] = "initial"
            dlg.destroy()

        def on_append():
            choice[0] = "append"
            dlg.destroy()

        def on_final():
            choice[0] = "final"
            dlg.destroy()

        dlg = tk.Toplevel(self.root)
        dlg.title("导出方式")
        dlg.transient(self.root)
        dlg.grab_set()
        ttk.Label(dlg, text="请选择导出方式：").pack(pady=(12, 8), padx=12)
        f = ttk.Frame(dlg, padding=8)
        f.pack(fill=tk.X)
        ttk.Button(f, text="初版 QCT", command=on_initial).pack(side=tk.LEFT, padx=4)
        ttk.Button(f, text="新增 Event", command=on_append).pack(side=tk.LEFT, padx=4)
        ttk.Button(f, text="终版 QCT", command=on_final).pack(side=tk.LEFT, padx=4)
        dlg.wait_window(dlg)
        return choice[0]

    def _export_qct(self):
        if not self.sdtm_rows and not self.adam_tfl_rows:
            messagebox.showwarning("提示", "请先导入 PDT 或 QCT 数据后再导出。")
            return
        export_mode = self._ask_export_qct_mode()
        if export_mode is None:
            return
        sdtm_to_write = self.sdtm_rows
        adam_to_write = self.adam_tfl_rows
        write_mode = export_mode
        if export_mode == "final":
            # 终版 QCT：仅保留 QC results description 非空的行
            sdtm_to_write = [r for r in self.sdtm_rows if len(r) > EDITABLE_COL_QC_DESC and (r[EDITABLE_COL_QC_DESC] or "").strip()]
            adam_to_write = [r for r in self.adam_tfl_rows if len(r) > EDITABLE_COL_QC_DESC and (r[EDITABLE_COL_QC_DESC] or "").strip()]
            write_mode = "append"
        if export_mode == "append":
            # 新增 Event：先选择要叠加的已有 QCT 文件，读入后与当前数据合并再保存
            initialdir = self._default_export_dir(prefer_qct=True)
            merge_path = filedialog.askopenfilename(
                title="选择要叠加的 QCT 文件（当前数据将追加到该文件）",
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx"), ("All", "*.*")],
                initialdir=initialdir,
            )
            if not merge_path:
                return
            try:
                existing_sdtm, existing_adam, _ = read_qct_workbook(merge_path)
                sdtm_to_write = existing_sdtm + self.sdtm_rows
                adam_to_write = existing_adam + self.adam_tfl_rows
            except Exception as e:
                messagebox.showerror("读取失败", f"无法读取要叠加的 QCT 文件：{e}")
                return
            initialdir = os.path.dirname(merge_path)
            initialfile = os.path.basename(merge_path)
        else:
            initialfile = self._default_export_name("QCT")
            initialdir = self._default_export_dir(prefer_qct=True)
        if export_mode == "final" and (not sdtm_to_write and not adam_to_write):
            messagebox.showwarning("提示", "终版 QCT：当前数据中无「QC results description」非空的行，无法导出。")
            return
        path = filedialog.asksaveasfilename(
            title="保存 QCT 文件",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("All", "*.*")],
            initialdir=initialdir,
            initialfile=initialfile,
        )
        if not path:
            return
        try:
            write_qct_workbook(
                sdtm_to_write, adam_to_write, path,
                pdt_path=self._pdt_path,
                event_value=self._event_var.get().strip(),
                export_mode=write_mode,
            )
            self._status.config(text=f"已导出: {path}", foreground="black")
            messagebox.showinfo("完成", f"QCT 已保存至:\n{path}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def _export_comments(self):
        if not self.sdtm_rows and not self.adam_tfl_rows:
            messagebox.showwarning("提示", "请先导入 PDT 或 QCT 数据后再导出。")
            return
        initialfile = self._default_export_name("Comments")
        initialdir = self._default_export_dir(prefer_qct=True)
        path = filedialog.asksaveasfilename(
            title="保存审阅意见（Comments）文件",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("All", "*.*")],
            initialdir=initialdir,
            initialfile=initialfile,
        )
        if not path:
            return
        try:
            write_comments_workbook(
                self.sdtm_rows, self.adam_tfl_rows, path,
                event_value=self._event_var.get().strip(),
            )
            self._status.config(text=f"已导出审阅意见: {path}", foreground="black")
            messagebox.showinfo("完成", f"审阅意见已保存至:\n{path}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def run(self):
        self.root.mainloop()


def main():
    app = QCTToolApp()
    app.run()


if __name__ == "__main__":
    main()
