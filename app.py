from __future__ import annotations

import queue
import sys
import threading
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageTk

from bill_splitter.core import BillSplitter, PauseController, ProgressUpdate, SplitResult, ValidationError
from bill_splitter.resources import export_resource, resource_path


APP_NAME = "账单拆分插入工具"
FILE_TYPES = [("Excel 文件", "*.xlsx *.xls"), ("Excel 2007+", "*.xlsx"), ("Excel 97-2003", "*.xls")]


class BillSplitterApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_NAME)
        self.geometry("780x610")
        self.minsize(720, 570)
        self.configure(bg="#f3f6f9")
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.pause_controller = PauseController()
        self.logo_image: ImageTk.PhotoImage | None = None

        self.source_var = tk.StringVar(value=str(resource_path("202601销售出库.xlsx")))
        self.template_var = tk.StringVar(value=str(resource_path("表头表尾格式.xlsx")))
        self.output_var = tk.StringVar(value=str(Path.cwd() / "拆分结果"))
        self.status_var = tk.StringVar(value="请选择文件和输出目录，然后点击“开始执行”。")
        self.current_var = tk.StringVar(value="当前客户：—")
        self.count_var = tk.StringVar(value="进度：0 / 0")

        self._configure_styles()
        self._set_icon()
        self._build_ui()
        self.after(100, self._poll_events)

    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 20, "bold"), foreground="#183153", background="#ffffff")
        style.configure("Sub.TLabel", font=("Microsoft YaHei UI", 10), foreground="#60738a", background="#ffffff")
        style.configure("Field.TLabel", font=("Microsoft YaHei UI", 10, "bold"), foreground="#26384a", background="#f3f6f9")
        style.configure("Status.TLabel", font=("Microsoft YaHei UI", 9), foreground="#44576a", background="#ffffff")
        style.configure("Primary.TButton", font=("Microsoft YaHei UI", 10, "bold"), padding=(18, 9))
        style.configure("Secondary.TButton", font=("Microsoft YaHei UI", 10), padding=(14, 8))
        style.configure("Horizontal.TProgressbar", thickness=15)

    def _set_icon(self) -> None:
        try:
            icon = resource_path("app.ico")
            if icon.exists():
                self.iconbitmap(default=str(icon))
            else:
                image = tk.PhotoImage(file=str(resource_path("logo.png")))
                self.iconphoto(True, image)
                self._window_icon = image
        except Exception:
            pass

    def _build_ui(self) -> None:
        header = tk.Frame(self, bg="#ffffff", height=118)
        header.pack(fill="x")
        header.pack_propagate(False)
        try:
            image = Image.open(resource_path("logo.png")).convert("RGBA")
            resampling = getattr(Image, "Resampling", Image).LANCZOS
            image.thumbnail((76, 76), resampling)
            self.logo_image = ImageTk.PhotoImage(image)
            tk.Label(header, image=self.logo_image, bg="#ffffff").pack(side="left", padx=(30, 18), pady=18)
        except Exception:
            pass
        title_box = tk.Frame(header, bg="#ffffff")
        title_box.pack(side="left", pady=21)
        ttk.Label(title_box, text=APP_NAME, style="Title.TLabel").pack(anchor="w")
        ttk.Label(title_box, text="按客户与月份拆分销售出库数据，并自动插入账单表头表尾", style="Sub.TLabel").pack(anchor="w", pady=(6, 0))

        content = tk.Frame(self, bg="#f3f6f9")
        content.pack(fill="both", expand=True, padx=28, pady=22)
        content.columnconfigure(1, weight=1)

        self._path_row(content, 0, "待拆分销售出库文件", self.source_var, self._choose_source)
        ttk.Button(content, text="下载销售出库模板", style="Secondary.TButton", command=lambda: self._download("202601销售出库.xlsx")).grid(row=1, column=1, sticky="w", pady=(5, 14))

        self._path_row(content, 2, "表头表尾格式文件", self.template_var, self._choose_template)
        ttk.Button(content, text="下载表头表尾模板", style="Secondary.TButton", command=lambda: self._download("表头表尾格式.xlsx")).grid(row=3, column=1, sticky="w", pady=(5, 14))

        self._path_row(content, 4, "输出目录", self.output_var, self._choose_output, directory=True)

        action_bar = tk.Frame(content, bg="#f3f6f9")
        action_bar.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(22, 16))
        self.start_button = ttk.Button(action_bar, text="开始执行", style="Primary.TButton", command=self._start)
        self.start_button.pack(side="left")
        self.pause_button = ttk.Button(action_bar, text="暂停", style="Primary.TButton", command=self._toggle_pause, state="disabled")
        self.pause_button.pack(side="left", padx=12)

        status_card = tk.Frame(content, bg="#ffffff", highlightbackground="#d9e2ec", highlightthickness=1)
        status_card.grid(row=6, column=0, columnspan=3, sticky="nsew")
        content.rowconfigure(6, weight=1)
        status_card.columnconfigure(0, weight=1)
        self.progress_bar = ttk.Progressbar(status_card, mode="determinate", maximum=100)
        self.progress_bar.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        ttk.Label(status_card, textvariable=self.count_var, style="Status.TLabel").grid(row=1, column=0, sticky="w", padx=20)
        ttk.Label(status_card, textvariable=self.current_var, style="Status.TLabel", wraplength=680).grid(row=2, column=0, sticky="w", padx=20, pady=(5, 0))
        ttk.Label(status_card, textvariable=self.status_var, style="Status.TLabel", wraplength=680).grid(row=3, column=0, sticky="nw", padx=20, pady=(5, 18))

    def _path_row(self, parent: tk.Widget, row: int, label: str, variable: tk.StringVar, command: object, directory: bool = False) -> None:
        ttk.Label(parent, text=label, style="Field.TLabel").grid(row=row, column=0, sticky="w", padx=(0, 14))
        entry = ttk.Entry(parent, textvariable=variable, font=("Microsoft YaHei UI", 10))
        entry.grid(row=row, column=1, sticky="ew", ipady=6)
        ttk.Button(parent, text="选择目录" if directory else "选择文件", command=command).grid(row=row, column=2, padx=(10, 0))

    def _choose_source(self) -> None:
        path = filedialog.askopenfilename(title="选择待拆分销售出库文件", filetypes=FILE_TYPES)
        if path:
            self.source_var.set(path)

    def _choose_template(self) -> None:
        path = filedialog.askopenfilename(title="选择表头表尾格式文件", filetypes=FILE_TYPES)
        if path:
            self.template_var.set(path)

    def _choose_output(self) -> None:
        path = filedialog.askdirectory(title="选择拆分结果输出目录")
        if path:
            self.output_var.set(path)

    def _download(self, name: str) -> None:
        destination = filedialog.asksaveasfilename(
            title=f"保存 {name}", initialfile=name, defaultextension=Path(name).suffix,
            filetypes=[("Excel 文件", f"*{Path(name).suffix}")],
        )
        if not destination:
            return
        try:
            export_resource(name, destination)
            messagebox.showinfo(APP_NAME, f"模板已保存到：\n{destination}")
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"模板保存失败：\n{exc}")

    def _start(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        source = self.source_var.get().strip().strip('"')
        template = self.template_var.get().strip().strip('"')
        output = self.output_var.get().strip().strip('"')
        if not source or not template or not output:
            messagebox.showwarning(APP_NAME, "请选择待拆分文件、表头表尾模板和输出目录。")
            return
        output_path = Path(output)
        if not output_path.exists():
            if not messagebox.askyesno(APP_NAME, f"输出目录不存在，是否创建？\n{output_path}"):
                return
            try:
                output_path.mkdir(parents=True)
            except OSError as exc:
                messagebox.showerror(APP_NAME, f"无法创建输出目录：\n{exc}")
                return

        self.pause_controller = PauseController()
        self.start_button.configure(state="disabled")
        self.pause_button.configure(state="normal", text="暂停")
        self.progress_bar.configure(value=0, maximum=100)
        self.count_var.set("进度：0 / 0")
        self.current_var.set("当前客户：—")
        self.status_var.set("正在启动……")
        splitter = BillSplitter(
            source, template, output_path,
            pause=self.pause_controller,
            progress=lambda update: self.events.put(("progress", update)),
            confirm_overwrite=self._confirm_overwrite_from_worker,
        )
        self.worker = threading.Thread(target=self._worker_run, args=(splitter,), daemon=True)
        self.worker.start()

    def _worker_run(self, splitter: BillSplitter) -> None:
        try:
            result = splitter.run()
            self.events.put(("done", result))
        except Exception as exc:
            self.events.put(("error", (exc, traceback.format_exc())))

    def _confirm_overwrite_from_worker(self, conflicts: list[Path]) -> bool:
        response: dict[str, bool] = {}
        ready = threading.Event()
        self.events.put(("confirm", (conflicts, response, ready)))
        ready.wait()
        return response.get("value", False)

    def _toggle_pause(self) -> None:
        if self.pause_controller.paused:
            self.pause_controller.resume()
            self.pause_button.configure(text="暂停")
            self.status_var.set("已继续执行。")
        else:
            self.pause_controller.pause()
            self.pause_button.configure(text="继续")
            self.status_var.set("正在安全暂停；当前写入步骤结束后停止……")

    def _poll_events(self) -> None:
        try:
            while True:
                event, payload = self.events.get_nowait()
                if event == "progress":
                    self._handle_progress(payload)  # type: ignore[arg-type]
                elif event == "confirm":
                    conflicts, response, ready = payload  # type: ignore[misc]
                    preview = "\n".join(path.name for path in conflicts[:8])
                    if len(conflicts) > 8:
                        preview += f"\n……另有 {len(conflicts) - 8} 个文件"
                    response["value"] = messagebox.askyesno(
                        APP_NAME,
                        f"输出目录中已有 {len(conflicts)} 个同名账单：\n\n{preview}\n\n是否覆盖这些文件？",
                    )
                    ready.set()
                elif event == "done":
                    self._handle_done(payload)  # type: ignore[arg-type]
                elif event == "error":
                    exc, detail = payload  # type: ignore[misc]
                    self._handle_error(exc, detail)
        except queue.Empty:
            pass
        self.after(100, self._poll_events)

    def _handle_progress(self, update: ProgressUpdate) -> None:
        if update.total:
            self.progress_bar.stop()
            self.progress_bar.configure(mode="determinate", maximum=update.total, value=update.completed)
            self.count_var.set(f"进度：{update.completed} / {update.total}")
        elif update.phase == "读取":
            self.progress_bar.configure(mode="indeterminate")
            self.progress_bar.start(12)
        if update.current_group:
            self.current_var.set(f"当前客户：{update.current_group}")
        self.status_var.set(f"{update.phase}：{update.message}")

    def _finish_ui(self) -> None:
        self.progress_bar.stop()
        self.progress_bar.configure(mode="determinate")
        self.start_button.configure(state="normal")
        self.pause_button.configure(state="disabled", text="暂停")
        self.pause_controller.resume()

    def _handle_done(self, result: SplitResult) -> None:
        self._finish_ui()
        self.progress_bar.configure(maximum=max(result.total, 1), value=result.total)
        self.count_var.set(f"进度：{result.total} / {result.total}")
        summary = (
            f"处理完成：成功 {result.succeeded} 个，失败 {result.failed} 个，"
            f"跳过空拆分键 {result.skipped_rows} 行。\n输出目录：{result.output_dir}"
        )
        if result.errors:
            summary += "\n\n失败明细：\n" + "\n".join(result.errors[:10])
            if len(result.errors) > 10:
                summary += f"\n……另有 {len(result.errors) - 10} 项"
        self.status_var.set(summary)
        if result.failed:
            messagebox.showwarning(APP_NAME, summary)
        else:
            messagebox.showinfo(APP_NAME, summary)

    def _handle_error(self, exc: Exception, detail: str) -> None:
        self._finish_ui()
        message = str(exc) if isinstance(exc, ValidationError) else f"执行失败：{exc}"
        self.status_var.set(message)
        messagebox.showerror(APP_NAME, message)
        if not isinstance(exc, ValidationError):
            print(detail, file=sys.stderr)

    def _on_close(self) -> None:
        if self.worker and self.worker.is_alive():
            if not messagebox.askyesno(APP_NAME, "任务仍在执行。关闭窗口会在当前步骤结束后终止界面，是否继续？"):
                return
            self.pause_controller.resume()
        self.destroy()


def main() -> None:
    app = BillSplitterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
