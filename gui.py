# -*- coding: utf-8 -*-
"""清泉出纳 · 考勤整理 · 图形启动器（tkinter 深色主题）

.bat 用 pythonw 跑本文件（无黑窗）。跑的子进程是本机 python + kaoqin.py。
风格模仿 collectors/tiktok-ads/gui.py。
"""

from __future__ import annotations

import os
import queue
import re
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

from tkinter import (
    BOTH,
    DISABLED,
    END,
    LEFT,
    NORMAL,
    RIGHT,
    WORD,
    X,
    Y,
    BooleanVar,
    Button,
    Canvas,
    Checkbutton,
    Frame,
    Label,
    Scrollbar,
    StringVar,
    Text,
    Tk,
    messagebox,
    ttk,
)

ROOT = Path(__file__).resolve().parent
LOGS = ROOT / "logs"
KAOQIN = ROOT / "kaoqin.py"
UPLOAD_SCRIPT = ROOT / "feishu-attendance-monthly.py"

# 找本机 python（优先和本进程同一个；gui 由 pythonw 启动，sys.executable 即 pythonw.exe）
PYW = sys.executable
# 子进程跑脚本用 python.exe（带 stdout），若 pythonw 同目录有 python.exe 就用它
_PY_DIR = Path(PYW).parent
PY = str(_PY_DIR / "python.exe") if (_PY_DIR / "python.exe").exists() else PYW


def _find_python() -> str:
    """子进程跑 kaoqin.py 用的 python：优先同目录 python.exe，否则 sys.executable。"""
    return PY


class App(Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("清泉出纳 · 考勤整理启动器")
        self.geometry("720x620")
        self.minsize(640, 540)
        self.configure(bg="#0f1419")
        self._proc: Optional[subprocess.Popen] = None
        self._q: queue.Queue[str] = queue.Queue()
        self._log_path: Optional[Path] = None
        self._build()
        self.after(100, self._drain_queue)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        LOGS.mkdir(parents=True, exist_ok=True)
        self._refresh_reports()

    # ---------------- UI ----------------
    def _build(self) -> None:
        # 头部
        header = Frame(self, bg="#1a222c", padx=16, pady=12)
        header.pack(fill="x")
        Label(
            header,
            text="清泉出纳 · 考勤整理启动器",
            font=("Microsoft YaHei UI", 16, "bold"),
            fg="#f8fafc",
            bg="#1a222c",
        ).pack(anchor="w")
        Label(
            header,
            text="选考勤报表 → 整理 → 出 3 张表结果 xlsx ·  可逐列比对现有结果表 ·  运行输出在底部",
            font=("Microsoft YaHei UI", 10),
            fg="#cbd5e1",
            bg="#1a222c",
        ).pack(anchor="w", pady=(4, 0))

        # footer（先 pack bottom，保证日志/状态固定可见）
        foot = Frame(self, bg="#0f1419", padx=16, pady=6)
        foot.pack(side="bottom", fill="x")
        Label(
            foot,
            text=f"{ROOT}  |  python: {Path(_find_python()).name}",
            font=("Segoe UI", 8),
            fg="#64748b",
            bg="#0f1419",
            anchor="w",
        ).pack(side=LEFT)

        # 状态栏
        status = Frame(self, bg="#0f1419", padx=16, pady=4)
        status.pack(side="bottom", fill="x")
        self.status_var = Label(
            status,
            text="就绪 · 选报表后点「整理」",
            font=("Microsoft YaHei UI", 10),
            fg="#e2e8f0",
            bg="#0f1419",
            anchor="w",
        )
        self.status_var.pack(side=LEFT, fill="x", expand=True)

        # 日志面板（固定高度，底部）
        log_frame = Frame(self, bg="#1a222c", padx=8, pady=6)
        log_frame.pack(side="bottom", fill=BOTH, expand=False)
        log_frame.configure(height=200)
        log_frame.pack_propagate(False)
        Label(
            log_frame,
            text="运行输出",
            font=("Microsoft YaHei UI", 10, "bold"),
            fg="#f1f5f9",
            bg="#1a222c",
            anchor="w",
        ).pack(fill="x", padx=4, pady=(2, 4))
        text_wrap = Frame(log_frame, bg="#1a222c")
        text_wrap.pack(fill=BOTH, expand=True, padx=4, pady=(0, 4))
        log_scroll = Scrollbar(text_wrap)
        log_scroll.pack(side=RIGHT, fill="y")
        self.log = Text(
            text_wrap,
            wrap=WORD,
            font=("Consolas", 9),
            bg="#0b1220",
            fg="#e2e8f0",
            insertbackground="#e2e8f0",
            relief="flat",
            yscrollcommand=log_scroll.set,
            height=8,
        )
        self.log.pack(side=LEFT, fill=BOTH, expand=True)
        self.log.configure(state=DISABLED)
        log_scroll.config(command=self.log.yview)

        # 上半区：输入 + 按钮区
        top = Frame(self, bg="#0f1419", padx=12, pady=8)
        top.pack(fill=BOTH, expand=True)

        # ---- 区块1: 考勤整理 ----
        sec1 = Frame(top, bg="#243044", padx=10, pady=5)
        sec1.pack(fill="x", padx=4, pady=(4, 8))
        Label(
            sec1,
            text="▍ 考勤整理",
            font=("Microsoft YaHei UI", 10, "bold"),
            fg="#f1f5f9",
            bg="#243044",
            anchor="w",
        ).pack(fill="x")

        row1 = Frame(sec1, bg="#243044")
        row1.pack(fill="x", pady=(4, 2))
        Label(
            row1, text="考勤报表:", font=("Microsoft YaHei UI", 10),
            fg="#cbd5e1", bg="#243044",
        ).pack(side=LEFT, padx=(0, 6))
        self.report_var = StringVar()
        self.cmb_report = ttk.Combobox(
            row1, textvariable=self.report_var, state="readonly", width=42,
        )
        self.cmb_report.pack(side=LEFT, padx=(0, 6))
        Button(
            row1, text="刷新", command=self._refresh_reports,
            font=("Microsoft YaHei UI", 9), fg="#ffffff", bg="#475569",
            activebackground="#475569", activeforeground="#ffffff",
            relief="flat", cursor="hand2", padx=10, pady=2,
        ).pack(side=LEFT)

        row2 = Frame(sec1, bg="#243044")
        row2.pack(fill="x", pady=(2, 6))
        self.check_var = BooleanVar(value=False)
        Checkbutton(
            row2, text="整理后逐列比对现有结果表", variable=self.check_var,
            font=("Microsoft YaHei UI", 10), fg="#e2e8f0", bg="#243044",
            selectcolor="#0b1220", activebackground="#243044", activeforeground="#e2e8f0",
            relief="flat", cursor="hand2",
        ).pack(side=LEFT)

        btns1 = Frame(sec1, bg="#243044")
        btns1.pack(fill="x", pady=(2, 4))
        Button(
            btns1, text="整理", command=self.run_run,
            font=("Microsoft YaHei UI", 10), fg="#ffffff", bg="#2563eb",
            activebackground="#2563eb", activeforeground="#ffffff",
            relief="flat", cursor="hand2", padx=12, pady=6,
        ).pack(side=LEFT, padx=(0, 6))
        Button(
            btns1, text="整理并比对", command=lambda: self.run_run(force_check=True),
            font=("Microsoft YaHei UI", 10), fg="#ffffff", bg="#0891b2",
            activebackground="#0891b2", activeforeground="#ffffff",
            relief="flat", cursor="hand2", padx=12, pady=6,
        ).pack(side=LEFT, padx=(0, 6))
        Button(
            btns1, text="上传到飞书", command=self.run_upload_feishu,
            font=("Microsoft YaHei UI", 10), fg="#ffffff", bg="#16a34a",
            activebackground="#16a34a", activeforeground="#ffffff",
            relief="flat", cursor="hand2", padx=12, pady=6,
        ).pack(side=LEFT)

        # ---- 区块2: 工具 ----
        sec2 = Frame(top, bg="#243044", padx=10, pady=5)
        sec2.pack(fill="x", padx=4, pady=(4, 8))
        Label(
            sec2,
            text="▍ 工具",
            font=("Microsoft YaHei UI", 10, "bold"),
            fg="#f1f5f9",
            bg="#243044",
            anchor="w",
        ).pack(fill="x")
        btns2 = Frame(sec2, bg="#243044")
        btns2.pack(fill="x", pady=(4, 4))
        for text, cmd, color in (
            ("打开结果目录", self.open_results, "#475569"),
            ("环境自检", self.doctor_only, "#f59e0b"),
            ("安装依赖 openpyxl", self.install_deps, "#f59e0b"),
            ("停止当前任务", self.stop_proc, "#dc2626"),
        ):
            Button(
                btns2, text=text, command=cmd,
                font=("Microsoft YaHei UI", 10), fg="#ffffff", bg=color,
                activebackground=color, activeforeground="#ffffff",
                relief="flat", cursor="hand2", padx=10, pady=6,
            ).pack(side=LEFT, padx=(0, 6))

    # ---------------- 日志 ----------------
    def _append(self, line: str) -> None:
        self.log.configure(state=NORMAL)
        self.log.insert(END, line + "\n")
        self.log.see(END)
        self.log.configure(state=DISABLED)

    def _drain_queue(self) -> None:
        try:
            while True:
                self._append(self._q.get_nowait())
        except queue.Empty:
            pass
        self.after(100, self._drain_queue)

    def _set_status(self, text: str) -> None:
        self.status_var.configure(text=text)

    def _new_run_log_path(self, title: str) -> Path:
        LOGS.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in title)[:40]
        return LOGS / f"gui-{stamp}-{safe}.log"

    def _busy(self) -> bool:
        return self._proc is not None and self._proc.poll() is None

    def _post(self, fn: Callable[[], None]) -> None:
        """主线程回调: 在 loop 就 after(0), 否则入队提示(不丢结束信号)。

        后台 worker 线程不能直接碰 tkinter; after() 在主线程不在 mainloop 时
        抛 RuntimeError。真实使用时 mainloop 一直在, 走 after 分支;
        退化分支(主循环已退出/测试中)只入队一条提示, 不跑 fn(避免在 worker 线程碰 messagebox)。
        """
        try:
            self.after(0, fn)
        except RuntimeError:
            self._q.put("[回调] 主循环已退出, 完成回调未执行")

    # ---------------- 文件扫描 ----------------
    def _refresh_reports(self) -> None:
        files = sorted(ROOT.glob("*考勤报表*.xlsx"))
        names = [f.name for f in files]
        self.cmb_report["values"] = names
        if names:
            self.cmb_report.current(0)
            self._set_status(f"就绪 · 找到 {len(names)} 个考勤报表，点「整理」")
        else:
            self.cmb_report.set("")
            self._set_status("未找到考勤报表，请把报表放到本目录后点「刷新」")

    def _find_existing_result(self) -> Optional[Path]:
        """现有结果表 = *考勤整理结果*.xlsx 排除脚本自己生成的 _校验，取最新。"""
        cands = [p for p in ROOT.glob("*考勤整理结果*.xlsx") if not p.stem.endswith("_校验")]
        if not cands:
            return None
        cands.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        return cands[0]

    def _find_latest_generated(self) -> Optional[Path]:
        cands = sorted(ROOT.glob("*考勤整理结果_校验*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
        return cands[0] if cands else None

    # ---------------- 运行子进程 ----------------
    def _run(self, title: str, cmd: list[str], confirm: str | None = None,
             cwd: Path | None = None, on_done: Optional[Callable[[int], None]] = None) -> None:
        if self._busy():
            messagebox.showwarning("进行中", "已有任务在运行，请先「停止当前任务」或等待结束。")
            return
        if confirm and not messagebox.askyesno("确认", confirm):
            return

        workdir = cwd or ROOT
        log_path = self._new_run_log_path(title)
        self._log_path = log_path
        header = [
            "",
            f"======== {datetime.now():%H:%M:%S} {title} ========",
            f"cwd={workdir}",
            " ".join(cmd),
            f"[log] {log_path}",
        ]
        for h in header:
            self._append(h)
        self._set_status(f"运行中：{title} · 日志 {log_path.name}")

        def worker() -> None:
            try:
                with open(log_path, "w", encoding="utf-8", newline="\n") as fp:
                    for h in header:
                        try:
                            fp.write(h + "\n")
                        except Exception:  # noqa: BLE001
                            pass
                    fp.flush()

                    def emit(s: str) -> None:
                        self._q.put(s)
                        try:
                            fp.write(s + "\n")
                            fp.flush()
                        except Exception:  # noqa: BLE001
                            pass

                    creation = 0
                    if sys.platform == "win32":
                        creation = getattr(subprocess, "CREATE_NO_WINDOW", 0)
                    self._proc = subprocess.Popen(
                        cmd,
                        cwd=str(workdir),
                        stdout=subprocess.PIPE,
                        stderr=subprocess.STDOUT,
                        text=True,
                        encoding="utf-8",
                        errors="replace",
                        creationflags=creation,
                    )
                    assert self._proc.stdout is not None
                    for line in self._proc.stdout:
                        emit(line.rstrip("\r\n"))
                    code = self._proc.wait()
                    emit(f"--- 结束 exit={code} ---")
                    emit(f"[log] saved: {log_path}")
                    if on_done:
                        self._post(lambda c=code: on_done(c))
                    else:
                        self._post(lambda c=code: self._on_task_done(title, c, log_path))
            except Exception as e:  # noqa: BLE001
                self._q.put(f"ERROR: {e}")
                self._post(lambda: self._set_status(f"失败：{title} · 见 {log_path.name}"))
            finally:
                self._proc = None

        threading.Thread(target=worker, daemon=True).start()

    def _on_task_done(self, title: str, code: int, log_path: Path) -> None:
        self._set_status(f"完成：{title} (exit={code}) · 日志 {log_path.name}")
        if code != 0:
            messagebox.showwarning(
                "任务失败",
                f"{title} 失败 (exit={code})。\n\n完整输出已写入：\n{log_path}\n\n可点「打开结果目录」查看。",
            )

    # ---------------- 业务动作 ----------------
    def run_run(self, force_check: bool = False) -> None:
        fn = self.report_var.get()
        if not fn:
            messagebox.showwarning("未选报表", "请先选择考勤报表（或点「刷新」）。")
            return
        src = ROOT / fn
        if not src.exists():
            messagebox.showerror("错误", f"报表文件不存在:\n{src}")
            return
        do_check = bool(self.check_var.get()) or force_check
        check_file: Optional[Path] = None
        if do_check:
            check_file = self._find_existing_result()
            if not check_file:
                messagebox.showinfo("比对", "未找到现有结果表（排除 _校验），跳过比对。")
        cmd = [PY, "-X", "utf8", str(KAOQIN), fn]
        if check_file:
            cmd += ["--check", check_file.name]
            title = "整理并比对"
            confirm = f"整理 {fn} 并与 {check_file.name} 逐列比对？"
        else:
            title = "整理"
            confirm = f"整理 {fn}？"

        def on_done(code: int) -> None:
            self._set_status(f"完成：{title} (exit={code})")
            if code != 0:
                messagebox.showwarning("失败", f"{title} 失败 (exit={code})，详见日志。")
                return
            result = self._find_latest_generated()
            if not result:
                self._q.put("[提示] 未找到生成的结果文件")
                return
            self._q.put(f"[完成] {result.name}")
            if messagebox.askyesno("完成", f"整理完成。\n结果文件: {result.name}\n\n打开结果文件？"):
                try:
                    os.startfile(str(result))  # type: ignore[attr-defined]
                except Exception as e:  # noqa: BLE001
                    messagebox.showerror("打开失败", str(e))

        self._run(title, cmd, confirm=confirm, on_done=on_done)

    def _detect_year_month(self, text: str) -> tuple[int, int]:
        """从文件名抓年月,与 feishu-attendance-monthly.py 的 detect_year_month 同口径。"""
        m = re.search(r"(\d{4})(\d{2})(\d{2})", text)
        if m:
            mo = int(m.group(2))
            if 1 <= mo <= 12:
                return int(m.group(1)), mo
        m = re.search(r"(\d{1,2})月", text)
        if m:
            mo = int(m.group(1))
            if 1 <= mo <= 12:
                ym = re.search(r"(\d{4})年", text)
                return (int(ym.group(1)) if ym else 2026), mo
        raise RuntimeError(f"无法从文件名识别月份: {text}")

    def run_upload_feishu(self) -> None:
        """上传当月考勤原始报表 + 整理结果到飞书(建/复用 5 张月表,清空后写入)。"""
        if not UPLOAD_SCRIPT.exists():
            messagebox.showerror("错误", f"上传脚本不存在:\n{UPLOAD_SCRIPT}")
            return
        fn = self.report_var.get()
        if not fn:
            messagebox.showwarning("未选报表", "请先选择考勤报表（或点「刷新」）。")
            return
        report = ROOT / fn
        if not report.exists():
            messagebox.showerror("错误", f"报表文件不存在:\n{report}")
            return
        result = self._find_latest_generated()
        if not result:
            messagebox.showwarning(
                "未找到整理结果",
                "没找到「*考勤整理结果_校验.xlsx」。\n请先点「整理」生成结果文件,再上传。",
            )
            return
        try:
            year, month = self._detect_year_month(report.name)
        except RuntimeError as e:  # noqa: BLE001
            messagebox.showerror("识别月份失败", str(e))
            return
        month_prefix = f"{month}月"
        confirm = (
            f"将 {year}年{month}月 考勤数据上传到飞书?\n\n"
            f"  报表: {report.name}\n"
            f"  整理结果: {result.name}\n\n"
            f"在飞书 base 里建/复用 5 张「{month_prefix}」前缀表\n"
            f"（考勤报表 / 打卡时间明细 / 考勤结果表 / 打卡异常说明 / 其他员工）,\n"
            f"清空后写入当月数据。历史月表不动。"
        )
        if not messagebox.askyesno("确认上传", confirm):
            return
        cmd = [PY, "-X", "utf8", str(UPLOAD_SCRIPT), str(report), str(result), "--apply"]
        title = "上传到飞书"

        def on_done(code: int) -> None:
            self._set_status(f"完成：{title} (exit={code})")
            if code != 0:
                messagebox.showwarning(
                    "上传失败",
                    f"{title} 失败 (exit={code})。\n完整日志已写入 logs 文件夹,详见运行输出。",
                )
                return
            if messagebox.askyesno(
                "完成",
                f"已上传 {year}年{month}月 考勤数据到飞书。\n\n"
                f"5 张「{month_prefix}」表已建/复用并写入当月数据。\n\n去飞书查看?",
            ):
                try:
                    os.startfile("https://dn86p8pyqr.feishu.cn/base/Xp7UbFxhRavG7OsxAP3cnAfXnJh")  # type: ignore[attr-defined]
                except Exception as e:  # noqa: BLE001
                    messagebox.showerror("打开失败", str(e))

        self._run(title, cmd, confirm=None, on_done=on_done)

    def open_results(self) -> None:
        os.startfile(str(ROOT))  # type: ignore[attr-defined]
        self._set_status("已打开结果目录")

    def doctor_only(self) -> None:
        cmd = [PY, "-c",
               "import sys; print('python', sys.version.split()[0]); "
               "import openpyxl; print('openpyxl', openpyxl.__version__)"]
        self._run("环境自检", cmd)

    def install_deps(self) -> None:
        cmd = [PY, "-m", "pip", "install", "openpyxl"]
        self._run("安装依赖", cmd, confirm="pip install openpyxl ？")

    # ---------------- 停止 / 关窗 ----------------
    def _kill_proc_tree(self, proc: subprocess.Popen) -> None:
        if proc.poll() is not None:
            return
        pid = proc.pid
        if sys.platform == "win32":
            subprocess.run(
                ["taskkill", "/PID", str(pid), "/T", "/F"],
                capture_output=True, text=True, encoding="utf-8", errors="replace",
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            try:
                proc.wait(timeout=8)
            except Exception:  # noqa: BLE001
                pass
            return
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except Exception:  # noqa: BLE001
            try:
                proc.kill()
            except Exception:  # noqa: BLE001
                pass

    def stop_proc(self) -> None:
        if not self._busy() or self._proc is None:
            messagebox.showinfo("停止", "当前没有运行中的任务。")
            return
        try:
            self._kill_proc_tree(self._proc)
            self._append("--- 已请求停止（含子进程树）---")
            self._set_status("已停止")
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("停止失败", str(e))

    def _on_close(self) -> None:
        if self._busy():
            if not messagebox.askyesno("退出", "任务仍在运行，确定退出？"):
                return
            try:
                if self._proc:
                    self._kill_proc_tree(self._proc)
            except Exception:  # noqa: BLE001
                pass
        self.destroy()


def main() -> None:
    if not KAOQIN.exists():
        # 无 GUI 也能报错（被 pythonw 启动时没控制台，弹窗报错）
        try:
            root = Tk(); root.withdraw()
            messagebox.showerror("错误", f"找不到 kaoqin.py:\n{KAOQIN}")
        except Exception:  # noqa: BLE001
            pass
        return
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
