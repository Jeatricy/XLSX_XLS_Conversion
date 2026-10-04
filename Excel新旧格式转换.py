"""
Excel 批量格式转换工具（xlsx ⇄ xls）
- 使用 CustomTkinter 现代化界面
- 强制使用 Excel COM 引擎（需安装 Microsoft Excel 和 pywin32）
- 默认包含子目录，保持目录结构，输出包含源根文件夹
- 同名冲突策略可选：跳过 / 覆盖 / 自动重命名
- 全局字体：微软雅黑（已放大至合适尺寸）
- 窗口图标：icon.ico（根目录下）
"""
import os
import traceback
from pathlib import Path
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox

# -------------------- 工具函数 --------------------
def ensure_dir(p: Path):
    p.mkdir(parents=True, exist_ok=True)

def iter_files(src_dir: Path, ext: str, recursive: bool):
    ext = ext.lower()
    if recursive:
        yield from (p for p in src_dir.rglob(f"*{ext}") if p.is_file())
    else:
        yield from (p for p in src_dir.glob(f"*{ext}") if p.is_file())

# -------------------- 转换引擎（仅 Excel COM） --------------------
def convert_with_excel_com(in_path: Path, out_path: Path):
    """使用 Microsoft Excel 进行转换，失败则抛出错误"""
    import win32com.client
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(in_path))
        out_ext = out_path.suffix.lower()
        if out_ext == ".xlsx":
            file_format = 51   # xlOpenXMLWorkbook
        elif out_ext == ".xls":
            file_format = 56   # xlExcel8
        else:
            raise ValueError(f"不支持的输出格式: {out_ext}")
        ensure_dir(out_path.parent)
        wb.SaveAs(str(out_path), FileFormat=file_format)
        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()

# -------------------- 路径规划 --------------------
def plan_output_path(in_path: Path, src_dir: Path, dst_dir: Path, out_ext: str):
    """保持目录结构 + 包含根文件夹"""
    rel = in_path.relative_to(src_dir)
    out_base = dst_dir / src_dir.name
    return (out_base / rel).with_suffix(out_ext)

def resolve_conflict(out_path: Path, conflict_mode: str) -> Path | None:
    if not out_path.exists():
        return out_path
    if conflict_mode == "overwrite":
        return out_path
    if conflict_mode == "skip":
        return None
    if conflict_mode == "rename":
        stem = out_path.stem
        suf = out_path.suffix
        parent = out_path.parent
        i = 1
        while True:
            cand = parent / f"{stem} ({i}){suf}"
            if not cand.exists():
                return cand
            i += 1
    raise ValueError("Invalid conflict_mode")

# -------------------- 批量转换 --------------------
def batch_convert(
    src_dir: Path, dst_dir: Path, mode: str,
    conflict_mode: str, stop_event: threading.Event, log_cb
):
    src_dir = src_dir.resolve()
    dst_dir = dst_dir.resolve()

    if mode == "xlsx2xls":
        in_ext, out_ext = ".xlsx", ".xls"
    elif mode == "xls2xlsx":
        in_ext, out_ext = ".xls", ".xlsx"
    else:
        raise ValueError("Invalid mode")

    files = sorted(iter_files(src_dir, in_ext, recursive=True))
    if not files:
        log_cb(f"未找到 {in_ext} 文件：{src_dir}\n")
        return 0, 0, 0

    ok = fail = skipped = 0
    log_cb(f"开始转换：{mode} | 文件数：{len(files)}\n")
    log_cb(f"冲突策略：{conflict_mode}\n\n")

    for i, in_path in enumerate(files, 1):
        if stop_event.is_set():
            log_cb("\n已停止：用户请求停止转换（将在当前文件完成后生效）。\n")
            break
        try:
            out_path = plan_output_path(in_path, src_dir, dst_dir, out_ext)
            final_out = resolve_conflict(out_path, conflict_mode)
            if final_out is None:
                skipped += 1
                log_cb(f"[{i}/{len(files)}] ↷ 跳过（已存在）：{out_path}\n")
                continue
            convert_with_excel_com(in_path, final_out)
            ok += 1
            log_cb(f"[{i}/{len(files)}] ✓ {in_path} -> {final_out}\n")
        except Exception as e:
            fail += 1
            log_cb(f"[{i}/{len(files)}] ✗ {in_path}\n  错误：{e}\n")

    log_cb(f"\n结束：成功 {ok}，失败 {fail}，跳过 {skipped}\n")
    return ok, fail, skipped

# -------------------- GUI (CustomTkinter) --------------------
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Excel 批量格式转换（xlsx ⇄ xls）")
        self.geometry("960x680")  # 稍大以容纳放大后的字体
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        # 窗口图标
        try:
            self.iconbitmap("icon.ico")
        except:
            pass  # 图标文件不存在则忽略

        # 统一字体定义（已放大）
        self.font_bold_20 = ctk.CTkFont("Microsoft YaHei", 20, "bold")
        self.font_bold_14 = ctk.CTkFont("Microsoft YaHei", 14, "bold")
        self.font_12 = ctk.CTkFont("Microsoft YaHei", 12)
        self.font_bold_12 = ctk.CTkFont("Microsoft YaHei", 12, "bold")

        # 变量
        self.src_var = ctk.StringVar()
        self.dst_var = ctk.StringVar()
        self.mode_var = ctk.StringVar(value="xlsx2xls")
        self.conflict_var = ctk.StringVar(value="skip")

        self.stop_event = threading.Event()
        self.worker = None

        self._build_ui()

    def _build_ui(self):
        # 主容器
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=20, pady=20)

        # 标题
        title = ctk.CTkLabel(main, text="Excel 批量转换", font=self.font_bold_20, text_color="#1e3a5f")
        title.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # ---------- 目录设置 ----------
        dir_frame = ctk.CTkFrame(main, corner_radius=10, fg_color="white")
        dir_frame.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        dir_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(dir_frame, text="📁 目录设置", font=self.font_bold_14, text_color="#1e3a5f").grid(
            row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(8, 4))

        ctk.CTkLabel(dir_frame, text="源目录：", width=80, font=self.font_12).grid(row=1, column=0, padx=12, pady=8, sticky="e")
        self.src_entry = ctk.CTkEntry(dir_frame, textvariable=self.src_var, placeholder_text="选择源文件夹...", font=self.font_12)
        self.src_entry.grid(row=1, column=1, padx=(8, 8), pady=8, sticky="ew")
        ctk.CTkButton(dir_frame, text="浏览…", width=80, font=self.font_bold_12, command=self.pick_src).grid(row=1, column=2, padx=12, pady=8)

        ctk.CTkLabel(dir_frame, text="输出目录：", width=80, font=self.font_12).grid(row=2, column=0, padx=12, pady=8, sticky="e")
        self.dst_entry = ctk.CTkEntry(dir_frame, textvariable=self.dst_var, placeholder_text="选择输出文件夹...", font=self.font_12)
        self.dst_entry.grid(row=2, column=1, padx=(8, 8), pady=8, sticky="ew")
        ctk.CTkButton(dir_frame, text="浏览…", width=80, font=self.font_bold_12, command=self.pick_dst).grid(row=2, column=2, padx=12, pady=8)

        # ---------- 转换配置 ----------
        cfg_frame = ctk.CTkFrame(main, corner_radius=10, fg_color="white")
        cfg_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        cfg_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(cfg_frame, text="⚙️ 转换配置", font=self.font_bold_14, text_color="#1e3a5f").grid(
            row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(8, 4))

        ctk.CTkLabel(cfg_frame, text="转换方向：", width=80, font=self.font_12).grid(row=1, column=0, padx=12, pady=8, sticky="e")
        dir_box = ctk.CTkFrame(cfg_frame, fg_color="transparent")
        dir_box.grid(row=1, column=1, sticky="w", padx=8, pady=8)
        ctk.CTkRadioButton(dir_box, text="xlsx → xls", variable=self.mode_var, value="xlsx2xls",
                           font=self.font_12, fg_color="#4a90e2").pack(side="left", padx=(0, 20))
        ctk.CTkRadioButton(dir_box, text="xls → xlsx", variable=self.mode_var, value="xls2xlsx",
                           font=self.font_12, fg_color="#4a90e2").pack(side="left")

        # ---------- 冲突策略 ----------
        conflict_frame = ctk.CTkFrame(main, corner_radius=10, fg_color="white")
        conflict_frame.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        conflict_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(conflict_frame, text="🔄 同名文件处理", font=self.font_bold_14, text_color="#1e3a5f").grid(
            row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(8, 4))

        ctk.CTkLabel(conflict_frame, text="策略：", width=80, font=self.font_12).grid(row=1, column=0, padx=12, pady=8, sticky="e")
        confl_box = ctk.CTkFrame(conflict_frame, fg_color="transparent")
        confl_box.grid(row=1, column=1, sticky="w", padx=8, pady=8)
        ctk.CTkRadioButton(confl_box, text="跳过", variable=self.conflict_var, value="skip",
                           font=self.font_12, fg_color="#4a90e2").pack(side="left", padx=(0, 20))
        ctk.CTkRadioButton(confl_box, text="覆盖", variable=self.conflict_var, value="overwrite",
                           font=self.font_12, fg_color="#4a90e2").pack(side="left", padx=(0, 20))
        ctk.CTkRadioButton(confl_box, text="自动重命名", variable=self.conflict_var, value="rename",
                           font=self.font_12, fg_color="#4a90e2").pack(side="left")

        # ---------- 操作按钮 ----------
        btn_frame = ctk.CTkFrame(main, fg_color="transparent")
        btn_frame.grid(row=4, column=0, columnspan=3, pady=(10, 12))
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(3, weight=1)

        self.start_btn = ctk.CTkButton(btn_frame, text="▶ 开始转换", fg_color="#27ae60", hover_color="#219a52",
                                       font=self.font_bold_12, command=self.run, height=36)
        self.start_btn.grid(row=0, column=1, padx=10)
        self.stop_btn = ctk.CTkButton(btn_frame, text="⏹ 停止", fg_color="#e74c3c", hover_color="#c0392b",
                                      font=self.font_bold_12, command=self.stop, state="disabled", height=36)
        self.stop_btn.grid(row=0, column=2, padx=10)

        # ---------- 日志 ----------
        log_frame = ctk.CTkFrame(main, corner_radius=10, fg_color="white")
        log_frame.grid(row=5, column=0, columnspan=3, sticky="nsew")
        log_frame.grid_rowconfigure(0, weight=1)
        log_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(log_frame, text="📋 运行日志", font=self.font_bold_14, text_color="#1e3a5f").grid(
            row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(8, 4))

        self.log_box = ctk.CTkTextbox(log_frame, wrap="word", font=self.font_12,
                                      fg_color="#f8faff", text_color="#333333", corner_radius=8)
        self.log_box.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=12, pady=(0, 12))

        # 权重使日志区可拉伸
        main.grid_rowconfigure(5, weight=1)
        main.grid_columnconfigure(0, weight=1)

    def _log(self, s: str):
        def append():
            self.log_box.insert("end", s)
            self.log_box.see("end")
        self.after(0, append)

    def pick_src(self):
        p = filedialog.askdirectory(title="选择源目录")
        if p:
            self.src_var.set(p)

    def pick_dst(self):
        p = filedialog.askdirectory(title="选择输出目录")
        if p:
            self.dst_var.set(p)

    def stop(self):
        self.stop_event.set()
        self._log("⏸ 正在请求停止…（当前文件完成后生效）\n")
        self.stop_btn.configure(state="disabled")

    def _reset_buttons(self):
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")

    def run(self):
        src = self.src_var.get().strip()
        dst = self.dst_var.get().strip()
        if not src or not Path(src).exists():
            messagebox.showerror("错误", "请选择有效的源目录")
            return
        if not dst:
            messagebox.showerror("错误", "请选择输出目录")
            return
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("提示", "转换正在进行中")
            return

        # 检查 Excel COM 可用性
        try:
            import win32com.client
            win32com.client.Dispatch("Excel.Application")
        except:
            messagebox.showerror("错误", "未检测到 Microsoft Excel 或 pywin32 未正确安装，无法使用 COM 引擎转换。")
            return

        self.stop_event.clear()
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")

        src_dir = Path(src)
        dst_dir = Path(dst)
        self._log("\n" + "=" * 40 + "\n")

        def task():
            try:
                ok, fail, skipped = batch_convert(
                    src_dir=src_dir, dst_dir=dst_dir,
                    mode=self.mode_var.get(),
                    conflict_mode=self.conflict_var.get(),
                    stop_event=self.stop_event,
                    log_cb=self._log
                )
                def done_msg():
                    messagebox.showinfo("结束", f"转换结束：成功 {ok}，失败 {fail}，跳过 {skipped}\n详情见日志。")
                self.after(0, done_msg)
            except Exception as e:
                tb = traceback.format_exc()
                self._log(tb + "\n")
                self.after(0, lambda: messagebox.showerror("异常", str(e)))
            finally:
                self.after(0, self._reset_buttons)

        self.worker = threading.Thread(target=task, daemon=True)
        self.worker.start()

if __name__ == "__main__":
    App().mainloop()