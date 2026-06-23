"""
converter_gui.py - 维普 Excel → ZD_JATS XML 转换器 (PyQt6 稳定版)

修改说明:
  - 移除了「从 Excel 读取」元数据的功能，界面仅保留核心转换逻辑。
"""

from __future__ import annotations

import sys
import io
import os
import threading
import traceback
from pathlib import Path
from typing import Optional

from contextlib import redirect_stderr, redirect_stdout

from PyQt6.QtCore import Qt, QObject, pyqtSignal
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLabel, QLineEdit, QPushButton, QCheckBox, 
    QTextEdit, QFileDialog, QMessageBox, QGroupBox
)
from PyQt6.QtGui import QFont

# 同目录导入转换核心（已移除未使用的元数据解析函数）
from converter import (
    DEFAULT_JOURNAL_META,
    DEFAULT_OUTPUT,
    DEFAULT_PDF_DIR,
    DEFAULT_XLSX,
    convert,
)


class LogSignaler(QObject):
    """用于在子线程中向主界面发送日志信号的类"""
    log_written = pyqtSignal(str)


class ConverterApp(QMainWindow):
    """转换器主窗口 (PyQt6 驱动)"""

    def __init__(self) -> None:
        super().__init__()
        os.environ.setdefault("TK_SILENCE_DEPRECATION", "1")
        
        self.signaler = LogSignaler()
        self.signaler.log_written.connect(self._log_msg)
        
        self.fields: dict[str, QLineEdit] = {}
        self._btn_convert: Optional[QPushButton] = None
        self._chk_pdf: Optional[QCheckBox] = None
        self._log: Optional[QTextEdit] = None

        self._init_ui()

    def _init_ui(self) -> None:
        self.setWindowTitle("维普 Excel → ZD_JATS XML 转换器")
        self.resize(850, 700)
        self.setMinimumSize(700, 600)

        # 全局主容器
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)

        # 1. 顶部标题区
        title_label = QLabel("维普 Excel → ZD_JATS XML")
        title_label.setFont(QFont("Arial", 18, QFont.Weight.Bold))
        main_layout.addWidget(title_label)

        hint_label = QLabel("选择表格与输出路径，填写期刊元数据后点击「开始转换」")
        hint_label.setFont(QFont("Arial", 12))
        hint_label.setStyleSheet("color: #666666;")
        main_layout.addWidget(hint_label)

        # 2. 路径设置区块
        paths_group = QGroupBox(" 路径设置 ")
        paths_group.setFont(QFont("Arial", 13, QFont.Weight.Bold))
        paths_layout = QGridLayout(paths_group)
        paths_layout.setContentsMargins(12, 16, 12, 12)
        paths_layout.setSpacing(10)

        self._add_path_row(paths_layout, 0, "Excel 表格:", "xlsx", "点击「浏览」选择维普 Excel 文件 (.xlsx)", self._browse_file)
        self._add_path_row(paths_layout, 1, "PDF 目录:", "pdf_dir", "PDF 根目录，支持子目录递归查找（可选）", self._browse_dir)
        self._add_path_row(paths_layout, 2, "输出目录:", "output_dir", "XML 输出目录，不存在将自动创建", self._browse_dir)

        self._chk_pdf = QCheckBox("同时复制 PDF 到输出目录")
        self._chk_pdf.setFont(QFont("Arial", 13))
        paths_layout.addWidget(self._chk_pdf, 3, 0, 1, 3)
        main_layout.addWidget(paths_group)

        # 3. 期刊元数据区块
        meta_group = QGroupBox(" 期刊元数据 ")
        meta_group.setFont(QFont("Arial", 13, QFont.Weight.Bold))
        meta_layout = QGridLayout(meta_group)
        meta_layout.setContentsMargins(12, 16, 12, 12)
        meta_layout.setSpacing(10)

        self._add_meta_field(meta_layout, 0, 0, "期刊中文名:", "journal_title", "机电工程")
        self._add_meta_field(meta_layout, 1, 0, "ISSN 号:", "issn", "1001-4551")
        self._add_meta_field(meta_layout, 0, 2, "国内刊号:", "cn", "33-1088/TH")
        self._add_meta_field(meta_layout, 1, 2, "出版社:", "publisher", "浙江大学")
        main_layout.addWidget(meta_group)

        # 4. 控制按钮区
        btn_bar = QHBoxLayout()
        self._btn_convert = QPushButton("开始转换")
        self._btn_convert.setFont(QFont("Arial", 13, QFont.Weight.Bold))
        self._btn_convert.setStyleSheet("background-color: #007aff; color: white; padding: 6px 14px;")
        self._btn_convert.clicked.connect(self._start_convert)
        btn_bar.addWidget(self._btn_convert)

        btn_exit = QPushButton("退出")
        btn_exit.setFont(QFont("Arial", 13))
        btn_exit.clicked.connect(self.close)
        btn_bar.addWidget(btn_exit)
        btn_bar.addStretch()
        main_layout.addLayout(btn_bar)

        # 5. 运行日志区块
        log_group = QGroupBox(" 运行日志 ")
        log_group.setFont(QFont("Arial", 13, QFont.Weight.Bold))
        log_layout = QVBoxLayout(log_group)
        
        self._log = QTextEdit()
        self._log.setFont(QFont("Menlo", 12))
        self._log.setReadOnly(True)
        self._log.setStyleSheet("background-color: #1e1e1e; color: #d4d4d4;")
        log_layout.addWidget(self._log)
        main_layout.addWidget(log_group)

        self._log_msg("就绪。输入框内灰色文字为占位提示，直接点击即可填写。")

    def _add_path_row(self, layout: QGridLayout, row: int, label_text: str, key: str, placeholder: str, callback: callable) -> None:
        lbl = QLabel(label_text)
        lbl.setFont(QFont("Arial", 13))
        lbl.setMinimumWidth(80)
        layout.addWidget(lbl, row, 0)

        edit = QLineEdit()
        edit.setFont(QFont("Arial", 13))
        edit.setPlaceholderText(placeholder)
        layout.addWidget(edit, row, 1)
        self.fields[key] = edit

        btn = QPushButton("浏览…")
        btn.setFont(QFont("Arial", 13))
        btn.clicked.connect(lambda: callback(key))
        layout.addWidget(btn, row, 2)

    def _add_meta_field(self, layout: QGridLayout, row: int, col_start: int, label_text: str, key: str, placeholder: str) -> None:
        lbl = QLabel(label_text)
        lbl.setFont(QFont("Arial", 13))
        lbl.setMinimumWidth(80)
        layout.addWidget(lbl, row, col_start)

        edit = QLineEdit()
        edit.setFont(QFont("Arial", 13))
        edit.setPlaceholderText(placeholder)
        layout.addWidget(edit, row, col_start + 1)
        self.fields[key] = edit

    # ---- 字段读写 ----

    def _get(self, key: str) -> str:
        return self.fields[key].text().strip()

    def _set(self, key: str, value: str) -> None:
        self.fields[key].setText(value)

    def _browse_file(self, key: str) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "选择 Excel 文件", "", "Excel 文件 (*.xlsx *.xls);;所有文件 (*)")
        if path:
            self._set(key, path)

    def _browse_dir(self, key: str) -> None:
        path = QFileDialog.getExistingDirectory(self, "选择目录")
        if path:
            self._set(key, path)

    def _log_msg(self, msg: str) -> None:
        if self._log:
            self._log.append(msg)
            self._log.ensureCursorVisible()

    def _journal_overrides(self) -> dict[str, str]:
        # 如果用户未填写，则自动降级使用核心文件里的默认配置项
        return {
            "title_zh": self._get("journal_title") or DEFAULT_JOURNAL_META.title_zh,
            "issn": self._get("issn") or DEFAULT_JOURNAL_META.issn,
            "cn": self._get("cn") or DEFAULT_JOURNAL_META.cn,
            "publisher": self._get("publisher") or DEFAULT_JOURNAL_META.publisher,
        }

    # ---- 业务逻辑 ----

    def _start_convert(self) -> None:
        xlsx = self._get("xlsx") or str(DEFAULT_XLSX)
        output = self._get("output_dir") or str(DEFAULT_OUTPUT)
        pdf_dir = None
        if self._chk_pdf and self._chk_pdf.isChecked():
            pdf_dir = self._get("pdf_dir") or str(DEFAULT_PDF_DIR)

        journal = self._journal_overrides()
        if self._btn_convert:
            self._btn_convert.setEnabled(False)
            
        self._log_msg("=" * 50)
        self._log_msg(f"[开始] {Path(xlsx).name}")

        def worker() -> None:
            buf = io.StringIO()
            count = 0
            ok = False
            try:
                with redirect_stdout(buf), redirect_stderr(buf):
                    count = convert(journal, xlsx, pdf_dir, output)
                ok = count > 0
            except Exception as exc:
                buf.write(f"[error] {exc}\n")

            output_text = buf.getvalue()

            for line in output_text.splitlines():
                if line.strip():
                    self.signaler.log_written.emit(line)

            def finish() -> None:
                if self._btn_convert:
                    self._btn_convert.setEnabled(True)
                if ok:
                    QMessageBox.information(
                        self, "完成",
                        f"转换成功，共生成 {count} 个 XML。\n输出目录:\n{Path(output)}",
                    )
                else:
                    QMessageBox.critical(self, "失败", "转换未完成，请查看日志。")

            from PyQt6.QtCore import QTimer
            QTimer.singleShot(0, finish)

        threading.Thread(target=worker, daemon=True).start()


def _crash_log_path() -> Path:
    if getattr(sys, "frozen", False):
        base = os.environ.get("LOCALAPPDATA") or str(Path.home())
        return Path(base) / "trans-converter-crash.log"
    return Path(__file__).resolve().parent / "trans-converter-crash.log"


def _write_crash_log(exc: BaseException) -> Path:
    log_path = _crash_log_path()
    log_path.write_text(
        "维普 Excel → ZD_JATS XML 转换器启动失败\n"
        f"Python: {sys.version}\n"
        f"Executable: {getattr(sys, 'executable', '')}\n\n"
        f"{traceback.format_exc()}",
        encoding="utf-8",
    )
    return log_path


def run_gui() -> None:
    app = QApplication(sys.argv)
    window = ConverterApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    try:
        run_gui()
    except Exception as exc:
        log_path = _write_crash_log(exc)
        try:
            from PyQt6.QtWidgets import QApplication, QMessageBox

            app = QApplication.instance() or QApplication(sys.argv)
            QMessageBox.critical(
                None,
                "启动失败",
                f"程序无法启动，错误已写入:\n{log_path}\n\n{exc}",
            )
        except Exception:
            print(f"[error] 程序无法启动，错误已写入: {log_path}", file=sys.stderr)
            print(exc, file=sys.stderr)
        sys.exit(1)