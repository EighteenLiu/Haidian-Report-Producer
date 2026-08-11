from __future__ import annotations

import argparse
import os
import time
import threading
from pathlib import Path

import customtkinter as ctk
from tkinter import filedialog, messagebox

from haidian_report.pipeline import generate_batch, generate_summary_only


ROOT = Path(__file__).resolve().parent
DEFAULT_DETAIL_TEMPLATE = ROOT / "街道案件明细表模板.xlsx"
DEFAULT_REPORT_TEMPLATE = ROOT / "街道环境建设管理工作运行情况分析报告模板.docx"
DEFAULT_INPUT_DIR = ROOT
DEFAULT_OUTPUT_DIR = ROOT / "out"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="街道分报告生成器")
    parser.add_argument("--input", nargs="*", help="输入文件或目录")
    parser.add_argument("--output", help="输出目录")
    parser.add_argument("--template-dir", help="模板目录", default=str(ROOT))
    parser.add_argument("--detail-template", help="案件明细表模板文件")
    parser.add_argument("--report-template", help="环境建设管理工作运行情况分析报告模板文件")
    return parser


class App(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("街道分报告生成器")
        self.geometry("1120x760")
        self.minsize(980, 680)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.input_files: list[Path] = []
        self.detail_template = Path(DEFAULT_DETAIL_TEMPLATE)
        self.report_template = Path(DEFAULT_REPORT_TEMPLATE)
        self.output_dir = DEFAULT_OUTPUT_DIR
        self._file_dialog_dir = DEFAULT_INPUT_DIR
        self._last_duration_text = ""

        self._build_ui()
        self._refresh_fields()

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=0)
        self.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self, corner_radius=0)
        header.grid(row=0, column=0, columnspan=2, sticky="ew", padx=12, pady=(12, 0))
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text="街道分报告生成器", font=ctk.CTkFont(size=20, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=16, pady=12
        )

        main = ctk.CTkFrame(self)
        main.grid(row=1, column=0, sticky="nsew", padx=(12, 6), pady=12)
        main.grid_columnconfigure(1, weight=1)
        main.grid_rowconfigure(3, weight=0)
        main.grid_rowconfigure(5, weight=1)

        self.detail_template_entry = self._add_entry_row(main, 0, "案件明细表模板", "选择文件", self.choose_detail_template)
        self.report_template_entry = self._add_entry_row(main, 1, "分析报告模板", "选择文件", self.choose_report_template)
        self.output_dir_entry = self._add_entry_row(main, 2, "输出目录", "选择目录", self.choose_output_dir)

        ctk.CTkLabel(main, text="已上传文件", anchor="w").grid(row=3, column=0, sticky="nw", padx=16, pady=(12, 0))
        self.files_box = ctk.CTkTextbox(main, height=150)
        self.files_box.grid(row=3, column=1, sticky="nsew", padx=(0, 16), pady=(12, 0))

        self.status_label = ctk.CTkLabel(main, text="等待操作", anchor="w")
        self.status_label.grid(row=4, column=0, columnspan=2, sticky="ew", padx=16, pady=(10, 0))

        self.log_box = ctk.CTkTextbox(main, height=220)
        self.log_box.grid(row=5, column=0, columnspan=2, sticky="nsew", padx=16, pady=(10, 16))

        side = ctk.CTkFrame(self)
        side.grid(row=1, column=1, sticky="ns", padx=(6, 12), pady=12)
        side.grid_columnconfigure(0, weight=1)

        self.add_files_button = ctk.CTkButton(side, text="添加文件", height=38, command=self.choose_source_files)
        self.add_files_button.grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 8))

        self.generate_button = ctk.CTkButton(side, text="生成", height=38, command=self.start_generation)
        self.generate_button.grid(row=1, column=0, sticky="ew", padx=14, pady=8)

        self.summary_button = ctk.CTkButton(side, text="仅生成表格提取表", height=38, command=self.start_summary_only)
        self.summary_button.grid(row=2, column=0, sticky="ew", padx=14, pady=8)

        self.open_output_button = ctk.CTkButton(side, text="打开输出目录", height=38, command=self.open_output_dir)
        self.open_output_button.grid(row=3, column=0, sticky="ew", padx=14, pady=8)

        self.clear_button = ctk.CTkButton(side, text="清空", height=38, command=self.clear_inputs)
        self.clear_button.grid(row=4, column=0, sticky="ew", padx=14, pady=8)

        self.progress = ctk.CTkProgressBar(side, progress_color="#2ea043")
        self.progress.grid(row=5, column=0, sticky="ew", padx=14, pady=(18, 6))
        self.progress.set(0)

        self.progress_label = ctk.CTkLabel(side, text="0%", anchor="w")
        self.progress_label.grid(row=6, column=0, sticky="ew", padx=14, pady=(0, 14))

    def _add_entry_row(
        self,
        parent: ctk.CTkFrame,
        row: int,
        label: str,
        button_text: str,
        command,
    ) -> ctk.CTkEntry:
        parent.grid_rowconfigure(row, weight=0)
        ctk.CTkLabel(parent, text=label, anchor="w", width=96).grid(
            row=row, column=0, sticky="w", padx=16, pady=(16 if row == 0 else 10, 0)
        )
        entry = ctk.CTkEntry(parent)
        entry.grid(row=row, column=1, sticky="ew", padx=(0, 16), pady=(16 if row == 0 else 10, 0))
        ctk.CTkButton(parent, text=button_text, width=110, command=command).grid(
            row=row, column=2, sticky="e", padx=(0, 16), pady=(16 if row == 0 else 10, 0)
        )
        return entry

    def _set_entry(self, entry: ctk.CTkEntry, value: str) -> None:
        entry.delete(0, "end")
        entry.insert(0, value)

    def _refresh_fields(self) -> None:
        self._set_entry(self.detail_template_entry, str(self.detail_template))
        self._set_entry(self.report_template_entry, str(self.report_template))
        self._set_entry(self.output_dir_entry, str(self.output_dir))

        self.files_box.configure(state="normal")
        self.files_box.delete("1.0", "end")
        if self.input_files:
            for file_path in self.input_files:
                self.files_box.insert("end", f"{file_path}\n")
        else:
            self.files_box.insert("end", "暂无已上传文件\n")
        self.files_box.configure(state="disabled")

    def _sync_paths_from_entries(self) -> None:
        detail_text = self.detail_template_entry.get().strip()
        report_text = self.report_template_entry.get().strip()
        output_text = self.output_dir_entry.get().strip()
        if detail_text:
            self.detail_template = Path(detail_text)
        if report_text:
            self.report_template = Path(report_text)
        if output_text:
            self.output_dir = Path(output_text)

    @staticmethod
    def _format_duration(seconds: float) -> str:
        total = max(0, int(round(seconds)))
        hours, rem = divmod(total, 3600)
        minutes, secs = divmod(rem, 60)
        if hours:
            return f"{hours}小时{minutes}分{secs}秒"
        if minutes:
            return f"{minutes}分{secs}秒"
        return f"{secs}秒"

    def log(self, text: str) -> None:
        self.log_box.configure(state="normal")
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.update_idletasks()

    def choose_detail_template(self) -> None:
        path = filedialog.askopenfilename(
            initialdir=str(self.detail_template.parent if self.detail_template.exists() else ROOT),
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.detail_template = Path(path)
            self._refresh_fields()

    def choose_report_template(self) -> None:
        path = filedialog.askopenfilename(
            initialdir=str(self.report_template.parent if self.report_template.exists() else ROOT),
            filetypes=[("Word files", "*.docx *.doc *.wps"), ("All files", "*.*")],
        )
        if path:
            self.report_template = Path(path)
            self._refresh_fields()

    def choose_output_dir(self) -> None:
        path = filedialog.askdirectory(initialdir=str(self.output_dir))
        if path:
            self.output_dir = Path(path)
            self._refresh_fields()

    def choose_source_files(self) -> None:
        files = filedialog.askopenfilenames(
            initialdir=str(self._file_dialog_dir),
            filetypes=[
                ("Supported files", "*.xlsx *.docx *.doc *.wps"),
                ("Office files", "*.xlsx *.docx *.doc *.wps"),
                ("All files", "*.*"),
            ],
        )
        if not files:
            return

        for item in files:
            path = Path(item)
            if path not in self.input_files:
                self.input_files.append(path)

        first_path = Path(files[0]).parent
        if first_path.exists():
            self._file_dialog_dir = first_path
        self._refresh_fields()

    def clear_inputs(self) -> None:
        self.input_files.clear()
        self.progress.set(0)
        self.progress_label.configure(text="0%")
        self.status_label.configure(text="已清空")
        self.log("已清空已上传文件")
        self._refresh_fields()

    def open_output_dir(self) -> None:
        self._sync_paths_from_entries()
        self.output_dir.mkdir(parents=True, exist_ok=True)
        os.startfile(self.output_dir)

    def _run_generation(self, summary_only: bool) -> None:
        self._sync_paths_from_entries()
        if not self.input_files:
            messagebox.showwarning("提示", "请先添加文件")
            return

        if not summary_only:
            if not self.detail_template.exists():
                messagebox.showwarning("提示", f"案件明细表模板不存在：{self.detail_template}")
                return
            if not self.report_template.exists():
                messagebox.showwarning("提示", f"分析报告模板不存在：{self.report_template}")
                return

        target_button = self.summary_button if summary_only else self.generate_button
        target_button.configure(state="disabled")
        started_at = time.perf_counter()
        self.progress.set(0)
        self.progress_label.configure(text="0%")
        self.status_label.configure(text="开始生成")
        self.log("开始生成任务" if not summary_only else "开始生成表格提取表")

        def worker() -> None:
            try:
                def on_progress(done: int, total: int, msg: str) -> None:
                    pct = done / total if total else 0
                    self.after(0, lambda: self.progress.set(pct))
                    self.after(0, lambda: self.progress_label.configure(text=f"{pct * 100:.0f}%"))
                    self.after(0, lambda: self.status_label.configure(text=msg))
                    self.after(0, lambda: self.log(msg))

                if summary_only:
                    results = generate_summary_only(list(self.input_files), self.output_dir, on_progress)
                else:
                    results = generate_batch(
                        list(self.input_files),
                        self.detail_template.parent,
                        self.output_dir,
                        on_progress,
                        self.detail_template,
                        self.report_template,
                    )
                self.after(0, lambda: self.progress.set(1))
                self.after(0, lambda: self.progress_label.configure(text="100%"))
                elapsed = self._format_duration(time.perf_counter() - started_at)
                self.after(0, lambda: self.status_label.configure(text=f"生成完成，用时 {elapsed}"))
                self.after(0, lambda: self.log(f"生成完成，共输出 {len(results)} 个文件，用时 {elapsed}"))
                self.after(0, lambda: messagebox.showinfo("完成", f"已生成 {len(results)} 个输出文件，用时 {elapsed}"))
            except Exception as exc:
                self.after(0, lambda: self.status_label.configure(text="生成失败"))
                self.after(0, lambda: self.log(f"失败: {exc}"))
                self.after(0, lambda: messagebox.showerror("错误", str(exc)))
            finally:
                self.after(0, lambda: target_button.configure(state="normal"))

        threading.Thread(target=worker, daemon=True).start()

    def start_generation(self) -> None:
        self._run_generation(summary_only=False)

    def start_summary_only(self) -> None:
        self._run_generation(summary_only=True)


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    if args.input and args.output:
        input_paths = [Path(x) for x in args.input]
        template_dir = Path(args.template_dir) if args.template_dir else ROOT
        detail_template = Path(args.detail_template) if args.detail_template else template_dir / DEFAULT_DETAIL_TEMPLATE.name
        report_template = Path(args.report_template) if args.report_template else template_dir / DEFAULT_REPORT_TEMPLATE.name
        output_dir = Path(args.output)
        generate_batch(input_paths, template_dir, output_dir, None, detail_template, report_template)
        print(str(output_dir))
        return

    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
