# -*- coding: utf-8 -*-
"""
QCT 小工具 GUI：导入 PDT / 导入 QCT → 选择 Event → 导出 QCT / 导出 Comments。
- 导出 QCT 时，两 Sheet 第一列（Event）使用当前选择的 Event。
- 新增 Event：可将当前数据追加到已有 QCT，并统一使用当前选择的 Event。
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from pdt_reader import read_and_clean_pdt

# ---------------------------------------------------------------------------
# 界面配色（微软风格）
# ---------------------------------------------------------------------------
COLORS = {
    "primary": "#0078d4",      # 微软蓝 - 主按钮（导入）
    "secondary": "#5e5e5e",    # 中灰色 - 次要
    "success": "#107c10",      # 绿色 - 成功/导出
    "warning": "#ffb900",      # 黄色 - 警告
    "error": "#d13438",        # 红色 - 错误
    "background": "#f5f5f5",  # 浅灰 - 窗口背景
    "surface": "#ffffff",      # 白色 - 卡片/面板背景
    "border": "#e0e0e0",      # 边框
    "text_primary": "#323130", # 主文字
    "text_secondary": "#605e5c",  # 次要文字
}

from qct_data import (
    build_qct_rows_from_pdt,
    read_qct_workbook,
    write_qct_workbook,
    write_comments_workbook,
    EVENT_OPTIONS_LIST,
    EDITABLE_COL_QC_DESC,
    EVENT_COL_INDEX,
)


# ---------------------------------------------------------------------------
# 主应用类
# ---------------------------------------------------------------------------

class QCTToolApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("QCT_Tools")
        self.root.minsize(520, 220)
        self.root.geometry("620x260")
        self.root.configure(bg=COLORS["background"])

        self.sdtm_rows = []
        self.adam_tfl_rows = []
        self._pdt_path = None
        self._qct_path = None  # 导入的 QCT 路径，用于导出时默认文件名

        self._build_ui()

    # ---------- UI 构建 ----------
    def _build_ui(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        # 导入按钮：主色（微软蓝）
        style.configure("Import.TButton", background=COLORS["primary"], foreground="white")
        style.map("Import.TButton", background=[("active", "#106ebe")], foreground=[("active", "white")])
        # 导出按钮：成功绿
        style.configure("Export.TButton", background=COLORS["success"], foreground="white")
        style.map("Export.TButton", background=[("active", "#0e6b0e")], foreground=[("active", "white")])
        # 弹窗内按钮：次要灰
        style.configure("Dialog.TButton", background=COLORS["secondary"], foreground="white")
        style.map("Dialog.TButton", background=[("active", "#4a4a4a")], foreground=[("active", "white")])
        # Add Event 按钮：与导入按钮同色（微软蓝）
        style.configure("AddEvent.TButton", background=COLORS["primary"], foreground="white")
        style.map("AddEvent.TButton", background=[("active", "#106ebe")], foreground=[("active", "white")])
        # 框架与标签
        style.configure("TFrame", background=COLORS["background"])
        style.configure("TLabel", background=COLORS["background"], foreground=COLORS["text_primary"])
        style.configure("TLabelframe", background=COLORS["surface"], bordercolor=COLORS["border"])
        style.configure("TLabelframe.Label", background=COLORS["surface"], foreground=COLORS["text_primary"])

        top = ttk.Frame(self.root, padding=10)
        top.pack(fill=tk.X)
        btn_row = ttk.Frame(top)
        btn_row.pack(fill=tk.X)
        ttk.Button(btn_row, text="📂 导入 PDT", command=self._import_pdt, style="Import.TButton").pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_row, text="📂 导入 QCT", command=self._import_qct, style="Import.TButton").pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_row, text="💾 导出 QCT", command=self._export_qct, style="Export.TButton").pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_row, text="📝 导出 Comments", command=self._export_comments, style="Export.TButton").pack(side=tk.LEFT, padx=4)
        self._status = ttk.Label(top, text="请先导入 PDT 文件", foreground=COLORS["text_secondary"])
        self._status.pack(fill=tk.X, pady=(8, 0), anchor=tk.W)

        # Event 选择区（卡片样式）：下拉框 + Add Event 按钮（可自定义添加 Event）
        edit_frame = ttk.LabelFrame(self.root, text="Event（导出 QCT 时将填充到两个 Sheet 的第一列）", padding=10)
        edit_frame.pack(fill=tk.X, padx=10, pady=10)
        ttk.Label(edit_frame, text="Event:").grid(row=0, column=0, sticky=tk.W, pady=2)
        event_row = ttk.Frame(edit_frame)
        event_row.grid(row=1, column=0, sticky=tk.W, pady=2)
        self._event_var = tk.StringVar(value="")
        self._combo_event = ttk.Combobox(
            event_row, textvariable=self._event_var, width=18,
            values=EVENT_OPTIONS_LIST, state="readonly",
        )
        self._combo_event.pack(side=tk.LEFT)
        ttk.Button(event_row, text="Add Event", command=self._add_event, width=10, style="AddEvent.TButton").pack(side=tk.LEFT, padx=(8, 0))

    # ---------- Event 相关：添加/输入 ----------
    def _ask_add_event_string(self):
        """弹出较大尺寸的输入框，标题「添加 Event」，提示「可在此输入新的Event」。返回输入的字符串或 None。"""
        result = [None]

        def on_ok():
            result[0] = entry.get().strip() if entry.get() else ""
            dlg.destroy()

        def on_cancel():
            dlg.destroy()

        dlg = tk.Toplevel(self.root)
        dlg.title("添加 Event")
        dlg.geometry("420x140")
        dlg.resizable(True, False)
        dlg.configure(bg=COLORS["background"])
        dlg.transient(self.root)
        dlg.grab_set()

        f = ttk.Frame(dlg, padding=20)
        f.pack(fill=tk.BOTH, expand=True)
        ttk.Label(f, text="可在此输入新的Event", font=("", 10)).pack(anchor=tk.W, pady=(0, 8))
        entry = ttk.Entry(f, width=42, font=("", 11))
        entry.pack(fill=tk.X, pady=(0, 16), ipady=4)
        entry.focus_set()
        btn_f = ttk.Frame(f)
        btn_f.pack(fill=tk.X)
        ttk.Button(btn_f, text="确定", command=on_ok, style="Import.TButton", width=8).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_f, text="取消", command=on_cancel, style="Dialog.TButton", width=8).pack(side=tk.LEFT)

        dlg.protocol("WM_DELETE_WINDOW", on_cancel)
        entry.bind("<Return>", lambda e: on_ok())
        self.root.wait_window(dlg)
        return result[0]

    def _add_event(self):
        """点击 Add Event 时弹出输入框，添加新 Event 并加入下拉列表、并选中。"""
        new_event = self._ask_add_event_string()
        if not new_event or not new_event.strip():
            return
        new_event = new_event.strip()
        if new_event in EVENT_OPTIONS_LIST:
            self._event_var.set(new_event)
            self._combo_event["values"] = list(EVENT_OPTIONS_LIST)
            messagebox.showinfo("添加 Event", f"「{new_event}」已在列表中，已为您选中。", parent=self.root)
            return
        EVENT_OPTIONS_LIST.append(new_event)
        self._combo_event["values"] = list(EVENT_OPTIONS_LIST)
        self._event_var.set(new_event)
        messagebox.showinfo("添加 Event", f"已添加「{new_event}」并选中。导出 QCT 时 List 表会包含该选项。", parent=self.root)

    # ---------- 导入：PDT / QCT ----------
    def _check_pdt_permission(self, path):
        """检查是否有访问该文件夹及编辑该 PDT 的权限。返回 (True, None) 或 (False, 错误信息)。"""
        try:
            dir_path = os.path.dirname(path)
            if not dir_path or not os.path.isdir(dir_path):
                return False, "所在文件夹不存在或无法访问。"
            try:
                os.listdir(dir_path)
            except PermissionError:
                return False, "没有访问该文件夹的权限。"
            except OSError as e:
                return False, f"无法访问该文件夹：{e}"
            if not os.path.isfile(path):
                return False, "该 PDT 文件不存在。"
            try:
                with open(path, "rb") as f:
                    f.read(1)
            except PermissionError:
                return False, "没有读取该 PDT 文件的权限。"
            except OSError as e:
                return False, f"无法读取该文件：{e}"
            try:
                with open(path, "r+b") as f:
                    pass
            except PermissionError:
                return False, "没有编辑该 PDT 文件的权限，无法导入。"
            except OSError as e:
                return False, f"无法编辑该文件：{e}"
            return True, None
        except Exception as e:
            return False, str(e)

    def _import_pdt(self):
        path = filedialog.askopenfilename(
            title="选择 PDT 文件",
            filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")],
        )
        if not path:
            return
        ok, err = self._check_pdt_permission(path)
        if not ok:
            messagebox.showerror("权限不足", err)
            return
        try:
            pdt_df = read_and_clean_pdt(path)
            event_val = self._event_var.get().strip()
            self.sdtm_rows, self.adam_tfl_rows = build_qct_rows_from_pdt(pdt_df, event_value=event_val)
            self._pdt_path = path
            self._qct_path = None
            self._status.config(
                text=f"已导入\n{os.path.basename(path)}  |  SDTM 行: {len(self.sdtm_rows)}  ADaM/TFL 行: {len(self.adam_tfl_rows)}",
                foreground=COLORS["text_primary"],
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
                foreground=COLORS["text_primary"],
            )
        except Exception as e:
            messagebox.showerror("导入 QCT 失败", str(e))

    # ---------- 导出：默认路径与文件名 ----------
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

    # ---------- 导出 QCT：方式选择与执行 ----------
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
        dlg.configure(bg=COLORS["background"])
        ttk.Label(dlg, text="请选择导出方式：").pack(pady=(14, 10), padx=14)
        f = ttk.Frame(dlg, padding=10)
        f.pack(fill=tk.X)
        ttk.Button(f, text="初版 QCT", command=on_initial, style="Dialog.TButton").pack(side=tk.LEFT, padx=4)
        ttk.Button(f, text="新增 Event", command=on_append, style="Dialog.TButton").pack(side=tk.LEFT, padx=4)
        ttk.Button(f, text="终版 QCT", command=on_final, style="Dialog.TButton").pack(side=tk.LEFT, padx=4)
        dlg.wait_window(dlg)
        return choice[0]

    def _export_qct(self):
        if not self.sdtm_rows and not self.adam_tfl_rows:
            messagebox.showwarning("提示", "请先导入 PDT 文件后再导出。")
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
            # 新增 Event：当前要追加的行统一使用界面所选 Event，再与已有 QCT 合并
            current_event = self._event_var.get().strip()
            for row in self.sdtm_rows:
                while len(row) <= EVENT_COL_INDEX:
                    row.append("")
                row[EVENT_COL_INDEX] = current_event
            for row in self.adam_tfl_rows:
                while len(row) <= EVENT_COL_INDEX:
                    row.append("")
                row[EVENT_COL_INDEX] = current_event
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
            self._status.config(text=f"已导出: {path}", foreground=COLORS["text_primary"])
            messagebox.showinfo("完成", f"QCT 已保存至:\n{path}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    # ---------- 导出 Comments ----------
    def _export_comments(self):
        if not self.sdtm_rows and not self.adam_tfl_rows:
            messagebox.showwarning("提示", "请先导入 PDT 文件后再导出。")
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
            self._status.config(text=f"已导出审阅意见: {path}", foreground=COLORS["text_primary"])
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
