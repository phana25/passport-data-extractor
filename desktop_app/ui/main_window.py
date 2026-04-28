from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, QSettings, QThread
from PySide6.QtGui import QKeySequence, QPixmap
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QFormLayout,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QStackedWidget,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from desktop_app import __version__
from desktop_app.services.extraction_worker import (
    ExtractionWorker,
    ScanResult,
    default_country_codes_path,
)
from desktop_app.services.history_store import HistoryStore, HistoryItem
from desktop_app.services.updater import CheckUpdateWorker, DownloadUpdateWorker, UpdaterService
from desktop_app.ui.preview import ImagePreview, UploadPreviewZone
from desktop_app.ui.styles import get_stylesheet
from passport_data_extractor import PassportDataExtractor


class CopyableTableWidget(QTableWidget):
    def keyPressEvent(self, event) -> None:
        if event.matches(QKeySequence.Copy):
            self._copy_selection_to_clipboard()
            return
        super().keyPressEvent(event)

    def _copy_selection_to_clipboard(self) -> None:
        indexes = self.selectedIndexes()
        if not indexes:
            return

        rows = sorted({index.row() for index in indexes})
        cols = sorted({index.column() for index in indexes})
        row_map = {row: i for i, row in enumerate(rows)}
        col_map = {col: i for i, col in enumerate(cols)}

        grid = [["" for _ in cols] for _ in rows]
        for index in indexes:
            item = self.item(index.row(), index.column())
            grid[row_map[index.row()]][col_map[index.column()]] = item.text() if item else ""

        text = "\n".join("\t".join(row) for row in grid)
        QApplication.clipboard().setText(text)


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Passport & Employee Data Verifier")
        self.setMinimumSize(1180, 720)
        self.resize(1200, 780)

        settings = QSettings("ByteLab", "PassportDataExtractor")
        self._theme = settings.value("theme", "light", type=str)
        self._apply_theme()

        self._passport_path: str | None = None
        self._card_path: str | None = None
        self._last_result: ScanResult | None = None
        self._extractor: PassportDataExtractor | None = None

        self._thread: QThread | None = None
        self._worker: ExtractionWorker | None = None
        self._history = HistoryStore()
        
        self._update_check_thread: QThread | None = None
        self._update_download_thread: QThread | None = None
        self._downloaded_update_path: str | None = None

        self._build_ui()
        self._check_for_updates()

    def keyPressEvent(self, event) -> None:
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            if hasattr(self, "stack") and self.stack.currentIndex() == 0:
                is_running = False
                if self._thread:
                    try:
                        is_running = self._thread.isRunning()
                    except RuntimeError:
                        self._thread = None
                if is_running:
                    pass
                elif hasattr(self, "btn_save_excel") and self.btn_save_excel.isEnabled():
                    self.btn_save_excel.click()
                elif hasattr(self, "btn_verify") and self.btn_verify.isEnabled():
                    if self._passport_path or self._card_path:
                        self.btn_verify.click()
        super().keyPressEvent(event)

    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("root")
        layout = QVBoxLayout(root)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        nav_bar = self._build_nav_bar()
        self.update_banner = self._build_update_banner()
        content = self._build_content()
        
        layout.addWidget(nav_bar)
        layout.addWidget(self.update_banner)
        layout.addWidget(content, 1)
        self.setCentralWidget(root)
        
    def _build_update_banner(self) -> QWidget:
        banner = QFrame()
        banner.setObjectName("update-banner")
        banner.setStyleSheet("background-color: #3b82f6; color: white;")
        banner.setFixedHeight(40)
        banner.setVisible(False)
        
        layout = QHBoxLayout(banner)
        layout.setContentsMargins(16, 4, 16, 4)
        
        self.update_label = QLabel("Checking for updates...")
        self.update_label.setStyleSheet("color: white; font-weight: bold;")
        
        self.update_progress = QProgressBar()
        self.update_progress.setRange(0, 100)
        self.update_progress.setFixedWidth(150)
        self.update_progress.setVisible(False)
        self.update_progress.setStyleSheet(
            "QProgressBar { background-color: #1d4ed8; color: white; border: none; text-align: center; border-radius: 4px; } "
            "QProgressBar::chunk { background-color: #60a5fa; border-radius: 4px; }"
        )
        
        self.btn_update_action = QPushButton("Download Update")
        self.btn_update_action.setStyleSheet("background-color: white; color: #3b82f6; font-weight: bold; border-radius: 4px; padding: 4px 12px;")
        self.btn_update_action.setVisible(False)
        self.btn_update_action.clicked.connect(self._on_update_action_clicked)
        
        btn_close = QPushButton("✕")
        btn_close.setFixedSize(24, 24)
        btn_close.setStyleSheet("background: transparent; color: white; font-weight: bold; font-size: 16px; border: none;")
        btn_close.clicked.connect(lambda: banner.setVisible(False))
        
        layout.addWidget(self.update_label)
        layout.addStretch()
        layout.addWidget(self.update_progress)
        layout.addWidget(self.btn_update_action)
        layout.addWidget(btn_close)
        
        return banner
        
    def _check_for_updates(self):
        if self._update_check_thread and self._update_check_thread.isRunning():
            return
            
        self._update_check_thread = QThread(self)
        self._update_checker = CheckUpdateWorker(__version__)
        self._update_checker.moveToThread(self._update_check_thread)
        
        self._update_check_thread.started.connect(self._update_checker.run)
        self._update_checker.update_available.connect(self._on_update_available)
        self._update_checker.no_update.connect(lambda: None)  # Do nothing silently
        self._update_checker.error.connect(lambda err: print(f"Update Check Error: {err}"))
        
        self._update_checker.update_available.connect(self._update_check_thread.quit)
        self._update_checker.no_update.connect(self._update_check_thread.quit)
        self._update_checker.error.connect(self._update_check_thread.quit)
        self._update_check_thread.finished.connect(self._update_check_thread.deleteLater)
        
        self._update_check_thread.start()
        
    def _on_update_available(self, version: str, download_url: str, notes: str):
        self._update_download_url = download_url
        self.update_label.setText(f"A new update (v{version}) is available!")
        self.btn_update_action.setText("Download")
        self.btn_update_action.setVisible(True)
        self.update_banner.setVisible(True)
        
    def _on_update_action_clicked(self):
        if self.btn_update_action.text() == "Download":
            self._start_update_download()
        elif self.btn_update_action.text() == "Restart to Install":
            self._install_and_restart()
            
    def _start_update_download(self):
        if not hasattr(self, '_update_download_url') or not self._update_download_url:
            return
            
        self.update_banner.setStyleSheet("background-color: #0f172a; color: white;")
        self.update_label.setText("Downloading update...")
        self.btn_update_action.setVisible(False)
        self.update_progress.setValue(0)
        self.update_progress.setVisible(True)
        
        self._update_download_thread = QThread(self)
        self._update_downloader = DownloadUpdateWorker(self._update_download_url)
        self._update_downloader.moveToThread(self._update_download_thread)
        
        self._update_download_thread.started.connect(self._update_downloader.run)
        self._update_downloader.progress.connect(self.update_progress.setValue)
        self._update_downloader.finished.connect(self._on_update_download_finished)
        self._update_downloader.error.connect(self._on_update_download_error)
        
        self._update_downloader.finished.connect(self._update_download_thread.quit)
        self._update_downloader.error.connect(self._update_download_thread.quit)
        self._update_download_thread.finished.connect(self._update_download_thread.deleteLater)
        
        self._update_download_thread.start()
        
    def _on_update_download_finished(self, dest_path: str):
        self._downloaded_update_path = dest_path
        self.update_progress.setVisible(False)
        self.update_banner.setStyleSheet("background-color: #10b981; color: white;")
        self.update_label.setText("Update download complete.")
        self.btn_update_action.setText("Restart to Install")
        self.btn_update_action.setVisible(True)
        
    def _on_update_download_error(self, err: str):
        self.update_progress.setVisible(False)
        self.update_banner.setStyleSheet("background-color: #ef4444; color: white;")
        self.update_label.setText(f"Update failed: {err}")
        self.btn_update_action.setVisible(False)
        
    def _install_and_restart(self):
        if self._downloaded_update_path:
            UpdaterService.install_and_restart(self._downloaded_update_path)
            QApplication.quit()

    def _apply_theme(self) -> None:
        """Apply theme stylesheet (light, dark, or system)."""
        QApplication.instance().setStyleSheet(get_stylesheet(self._theme))
        if hasattr(self, "history_tabs"):
            self._refresh_history()

    def _on_theme_changed(self, text: str) -> None:
        theme_map = {"Light": "light", "Dark": "dark", "System": "system"}
        self._theme = theme_map.get(text, "light")
        QSettings("ByteLab", "PassportDataExtractor").setValue("theme", self._theme)
        self._apply_theme()

    def _build_nav_bar(self) -> QWidget:
        bar = QFrame()
        bar.setObjectName("nav-bar")
        bar.setFixedHeight(56)
        bar.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        layout = QHBoxLayout(bar)
        layout.setContentsMargins(24, 12, 24, 12)
        layout.setSpacing(16)
        layout.setAlignment(Qt.AlignVCenter)

        title = QLabel(f"v{__version__}")
        title.setObjectName("sidebar-title")
        layout.addWidget(title, 0, Qt.AlignVCenter)

        self.btn_nav_scan = QPushButton("Document Scan")
        self.btn_nav_history = QPushButton("Data List")
        for btn in (self.btn_nav_scan, self.btn_nav_history):
            btn.setCheckable(True)
            btn.setFixedHeight(32)
            btn.setObjectName("nav-btn")
            
        self.btn_nav_scan.setChecked(True)
        self.btn_nav_scan.clicked.connect(lambda: self._on_nav_changed(0))
        self.btn_nav_history.clicked.connect(lambda: self._on_nav_changed(1))
        layout.addWidget(self.btn_nav_scan, 0, Qt.AlignVCenter)
        layout.addWidget(self.btn_nav_history, 0, Qt.AlignVCenter)

        layout.addStretch(1)

        theme_lbl = QLabel("Theme")
        theme_lbl.setObjectName("theme-label")
        theme_lbl.setAlignment(Qt.AlignVCenter | Qt.AlignRight)
        
        self.theme_combo = QComboBox()
        self.theme_combo.setObjectName("theme-combo")
        self.theme_combo.addItems(["Light", "Dark", "System"])
        self.theme_combo.setCurrentText(
            {"light": "Light", "dark": "Dark", "system": "System"}.get(self._theme, "Light")
        )
        self.theme_combo.setFixedWidth(88)
        self.theme_combo.setFixedHeight(32)
        self.theme_combo.currentTextChanged.connect(self._on_theme_changed)
        layout.addWidget(theme_lbl, 0, Qt.AlignVCenter)
        layout.addWidget(self.theme_combo, 0, Qt.AlignVCenter)

        return bar

    def _build_content(self) -> QWidget:
        wrap = QWidget()
        layout = QVBoxLayout(wrap)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(12)

        self.stack = QStackedWidget()
        self.stack.addWidget(self._build_scan_page())
        self.stack.addWidget(self._build_history_page())
        layout.addWidget(self.stack, 1)

        self._on_nav_changed(0)
        return wrap

    def _build_scan_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setSpacing(16)

        # Upload + preview section — two zones; scan works with 1 or both
        preview_section = QWidget()
        preview_layout = QHBoxLayout(preview_section)
        preview_layout.setContentsMargins(0, 0, 0, 0)
        preview_layout.setSpacing(12)

        self.passport_preview = UploadPreviewZone(
            "Passport",
            self._pick_passport,
            on_clear=lambda: setattr(self, "_passport_path", None),
        )
        self.card_preview = UploadPreviewZone(
            "Employee Card",
            self._pick_card,
            on_clear=lambda: setattr(self, "_card_path", None),
        )

        preview_layout.addWidget(self.passport_preview, 1)
        preview_layout.addWidget(self.card_preview, 1)
        preview_section.setMaximumHeight(260)
        layout.addWidget(preview_section, 0)

        # Data section — no scroll, compact layout
        layout.addWidget(self._build_data_section(), 1)

        # Actions
        action_row = QHBoxLayout()
        action_row.setSpacing(12)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setTextVisible(False)
        self.progress.setFixedHeight(12)
        self.progress.setVisible(False)

        self.btn_verify = QPushButton("Verify Data")
        self.btn_verify.setMinimumWidth(140)
        self.btn_verify.setStyleSheet(
            "QPushButton { background: #3b82f6; color: white; } "
            "QPushButton:hover { background: #2563eb; } "
            "QPushButton:disabled { background: #94a3b8; color: #e2e8f0; }"
        )
        self.btn_verify.clicked.connect(self._start_scan)

        self.btn_save_excel = QPushButton("Save Data to List")
        self.btn_save_excel.setMinimumWidth(140)
        self.btn_save_excel.setEnabled(False)
        self.btn_save_excel.setStyleSheet(
            "QPushButton { background: #10b981; color: white; } "
            "QPushButton:hover { background: #059669; } "
            "QPushButton:disabled { background: #94a3b8; color: #e2e8f0; }"
        )
        self.btn_save_excel.clicked.connect(self._save_data)

        self.btn_clear = QPushButton("Clear")
        self.btn_clear.setStyleSheet(
            "QPushButton { background: #e2e8f0; color: #475569; } "
            "QPushButton:hover { background: #cbd5e1; }"
        )
        self.btn_clear.clicked.connect(self._clear)

        self.status_label = QLabel("Upload at least one image (passport and/or employee card), then click Verify.")
        self.status_label.setStyleSheet("color: #64748b; font-size: 13px;")

        action_row.addWidget(self.btn_verify)
        action_row.addWidget(self.btn_save_excel)
        action_row.addWidget(self.btn_clear)
        action_row.addSpacing(16)
        action_row.addWidget(self.progress, 1)
        action_row.addWidget(self.status_label, 2)
        layout.addLayout(action_row)

        return page

    def _build_data_section(self) -> QWidget:
        wrap = QFrame()
        wrap.setObjectName("data-card")
        main = QVBoxLayout(wrap)
        main.setContentsMargins(0, 0, 0, 0)
        main.setSpacing(16)

        # Two balanced columns side by side
        cols = QHBoxLayout()
        cols.setSpacing(16)

        def make_section(title: str, fields: list[str], out_map: dict, columns: int = 1) -> QWidget:
            box = QFrame()
            box.setObjectName("data-section-box")
            lay = QVBoxLayout(box)
            lay.setSpacing(12)
            lay.setContentsMargins(16, 16, 16, 16)
            
            header = QLabel(title)
            header.setObjectName("data-section-header")
            lay.addWidget(header)
            
            grid = QGridLayout()
            grid.setSpacing(8)
            
            for i, name in enumerate(fields):
                field_container = QWidget()
                field_lay = QVBoxLayout(field_container)
                field_lay.setSpacing(2)
                field_lay.setContentsMargins(0, 0, 0, 0)

                label_row = QHBoxLayout()
                label_row.setContentsMargins(0, 0, 0, 0)
                label_row.setSpacing(4)

                lbl = QLabel(name)
                lbl.setObjectName("data-field-label")
                label_row.addWidget(lbl)
                label_row.addStretch()

                val = QLineEdit()
                val.setObjectName("data-field")
                val.setPlaceholderText("—")
                val.setMinimumHeight(30)

                if name == "NAME 02":
                    self.btn_name02_recalc = QPushButton("↻")
                    self.btn_name02_recalc.setFixedSize(16, 16)
                    self.btn_name02_recalc.setCursor(Qt.PointingHandCursor)
                    self.btn_name02_recalc.setToolTip("Recalculate NAME 02 from SURNAME and GSURNAME")
                    self.btn_name02_recalc.setStyleSheet(
                        "QPushButton { background: transparent; color: #94a3b8; font-weight: bold; border: none; padding: 0; } "
                        "QPushButton:hover { color: #f8fafc; }"
                    )
                    self.btn_name02_recalc.clicked.connect(self._recalculate_name02_from_fields)
                    label_row.addWidget(self.btn_name02_recalc, 0, Qt.AlignVCenter)

                field_lay.addLayout(label_row)
                if name == "NAME 02":
                    field_lay.addWidget(val)
                else:
                    field_lay.addWidget(val)
                
                row = i // columns
                col = i % columns
                grid.addWidget(field_container, row, col)
                out_map[name] = val
                
            lay.addLayout(grid)
            lay.addStretch()
            return box

        passport_fields = [
            "SURNAME", "GSURNAME", "PASSPORT",
            "BD1", "BD2", "BD3",
            "ISS1", "ISS2", "ISS3",
            "ED1", "ED2", "ED3",
            "NASTIONALTY", "NAME 02", "Gender",
        ]
        card_fields = [
            "CARD NUMBER",
            "COMPANY CARD", "POSITOIN CARD",
            "DC1", "DC2", "DC3",
        ]
        other_fields = [
            "PHONE", "COMPANY", "POSTION",
            "D01", "D02", "D03",
        ]

        self.passport_out = {}
        self.card_out = {}
        self.other_out = {}
        p_section = make_section("Passport", passport_fields, self.passport_out, columns=3)
        c_section = make_section("Employee Card", card_fields, self.card_out, columns=3)
        o_section = make_section("Other", other_fields, self.other_out, columns=3)
        
        cols.addWidget(p_section, 1)
        
        right_col = QVBoxLayout()
        right_col.setContentsMargins(0, 0, 0, 0)
        right_col.setSpacing(16)
        right_col.addWidget(c_section)
        right_col.addWidget(o_section)
        right_col.addStretch()
        cols.addLayout(right_col, 1)
        
        main.addLayout(cols)

        # Options bar — full width, subtle
        opts = QFrame()
        opts.setObjectName("opts-bar")
        opts_layout = QHBoxLayout(opts)
        opts_layout.setContentsMargins(0, 0, 0, 0)
        opts_layout.setSpacing(24)
        v_lbl = QLabel("Validity")
        v_lbl.setObjectName("opts-label")
        opts_layout.addWidget(v_lbl)
        self.validity_combo = QComboBox()
        self.validity_combo.setObjectName("validity-combo")
        self.validity_combo.addItems(["12M", "6M", "3M", "1M", "M"])
        self.validity_combo.setFixedWidth(80)
        opts_layout.addWidget(self.validity_combo)
        opts_layout.addStretch()
        main.addWidget(opts)

        return wrap

    def _build_history_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setSpacing(16)

        top = QHBoxLayout()
        btn_refresh = QPushButton("Refresh")
        btn_refresh.setStyleSheet("background: #e2e8f0; color: #475569;")
        btn_refresh.clicked.connect(self._refresh_history)
        btn_export_all = QPushButton("Export Current Tab")
        btn_export_all.setStyleSheet("background: #10b981; color: white;")
        btn_export_all.clicked.connect(self._export_current_tab)
        btn_clear = QPushButton("Clear Data List")
        btn_clear.setStyleSheet("background: #fee2e2; color: #dc2626;")
        btn_clear.clicked.connect(self._clear_history)
        top.addWidget(btn_refresh)
        top.addWidget(btn_export_all)
        top.addWidget(btn_clear)
        top.addStretch(1)
        layout.addLayout(top)

        from passport_data_extractor import PassportDataExtractor
        self.excel_headers = PassportDataExtractor.EXCEL_HEADERS

        self.history_tabs = QTabWidget()
        self.history_tabs.setTabPosition(QTabWidget.South)
        self.history_tabs.setStyleSheet("""
            QTabWidget::pane { border: none; background: transparent; }
            QTabWidget::tab-bar { alignment: left; }
            QTabBar::tab { background: transparent; padding: 6px 14px; margin-right: 2px; color: #94a3b8; border-bottom: 2px solid transparent; }
            QTabBar::tab:selected { font-weight: bold; color: #3b82f6; border-bottom: 2px solid #3b82f6; }
            QTabBar::tab:hover:!selected { color: #3b82f6; }
        """)
        
        layout.addWidget(self.history_tabs, 1)

        return page

    def _on_nav_changed(self, idx: int) -> None:
        self.btn_nav_scan.setChecked(idx == 0)
        self.btn_nav_history.setChecked(idx == 1)
        if hasattr(self, "stack"):
            self.stack.setCurrentIndex(idx)
        if idx == 1:
            self._refresh_history()

    def _refresh_history(self, target_export_date: str | None = None) -> None:
        if not hasattr(self, "history_tabs"):
            return

        # Save current tab selection by title
        current_idx = self.history_tabs.currentIndex()
        current_title_base = self.history_tabs.tabText(current_idx) if current_idx >= 0 else None

        items = self._history.load()
        
        # Disabling signals to prevent recursive calls during rebuild
        self.history_tabs.blockSignals(True)
        while self.history_tabs.count() > 0:
            self.history_tabs.removeTab(0)
            
        if not items:
            self.history_tabs.blockSignals(False)
            return

        from PySide6.QtGui import QColor
        from PySide6.QtCore import Qt
        
        color_exported = QColor("#94a3b8")
        
        # Group items by export date
        groups = {}
        for it in items:
            is_exported = getattr(it, "exported", False)
            if not is_exported:
                group_key = "NEW"
            else:
                exp_date = getattr(it, "export_date", None)
                group_key = f"Exported: {exp_date}" if exp_date else "Exported (Old)"
                
            if group_key not in groups:
                groups[group_key] = []
            groups[group_key].append(it)
            
        def create_tab_content_for_group(group_items, is_exported_group):
            page = QWidget()
            lay = QVBoxLayout(page)
            lay.setContentsMargins(0, 0, 0, 0)
            # Theme-aware colors
            current_theme = getattr(self, "theme_combo", None)
            is_dark = (current_theme.currentText() == "Dark") if current_theme else (self._theme == 'dark')
            
            bg_color = "#1e293b" if is_dark else "#ffffff"
            text_color = "#f1f5f9" if is_dark else "#0f172a"
            grid_color = "#334155" if is_dark else "#e2e8f0"
            header_bg = "#334155" if is_dark else "#f1f5f9"
            header_text = "#94a3b8" if is_dark else "#475569"
            stats_bg = "#0f172a" if is_dark else "#ffffff"
            stats_text_color = "#94a3b8" if is_dark else "#475569"
            
            # Generate statistics
            total = len(group_items)
            males = sum(1 for item in group_items if (item.combined or {}).get("M") == "X")
            females = sum(1 for item in group_items if (item.combined or {}).get("F") == "X")
            stats_text = f"Records: {total}   |   Male: {males}   |   Female: {females}"
            
            stats_lbl = QLabel(stats_text)
            stats_lbl.setStyleSheet(f"color: {stats_text_color}; font-size: 13px; font-weight: bold; background: {stats_bg}; padding: 6px 12px; border-bottom: 1px solid {grid_color};")
            lay.addWidget(stats_lbl)
            
            # Prepare display headers (M_VAL -> M)
            display_headers = [h if h != 'M_VAL' else 'M' for h in self.excel_headers]
            
            table = CopyableTableWidget()
            table.setColumnCount(len(display_headers))
            table.setHorizontalHeaderLabels(display_headers)
            table.verticalHeader().setVisible(True)
            table.setEditTriggers(QAbstractItemView.NoEditTriggers)
            table.setSelectionMode(QAbstractItemView.ExtendedSelection)
            table.setSelectionBehavior(QAbstractItemView.SelectItems)
            table.setAlternatingRowColors(True)
            table.setCornerButtonEnabled(True)
            table.horizontalHeader().setSectionsClickable(True)
            table.verticalHeader().setSectionsClickable(True)
            table.horizontalHeader().sectionClicked.connect(table.selectColumn)
            table.verticalHeader().sectionClicked.connect(table.selectRow)
            table.setStyleSheet(f"""
                QTableWidget {{ background: {bg_color}; color: {text_color}; border: none; gridline-color: {grid_color}; }}
                QHeaderView {{ background-color: {header_bg}; border: none; }}
                QHeaderView::section {{ background-color: {header_bg}; color: {header_text}; padding: 4px; border: 1px solid {grid_color}; font-weight: bold; border-top: none; }}
                QTableWidget::item {{ border-bottom: 1px solid {grid_color}; }}
            """)
            
            table.setRowCount(len(group_items))
            for row_idx, it in enumerate(group_items):
                combined = it.combined or {}
                for col_idx, header in enumerate(self.excel_headers):
                    val = combined.get(header, "")
                    if val == "Not Found" or val is None:
                        val = ""
                    item = QTableWidgetItem(str(val))
                    item.setTextAlignment(Qt.AlignCenter)
                    if is_exported_group:
                        item.setForeground(color_exported)
                    table.setItem(row_idx, col_idx, item)
            
            table.resizeColumnsToContents()
            lay.addWidget(table, 1)
            
            return page

        sorted_keys = sorted([k for k in groups.keys() if k != "NEW"])
        
        for key in sorted_keys:
            page = create_tab_content_for_group(groups[key], True)
            title_base = key.replace("Exported: ", "")
            self.history_tabs.addTab(page, title_base)
            
        if "NEW" in groups:
            page = create_tab_content_for_group(groups["NEW"], False)
            title_base = "New Data (Ready)"
            self.history_tabs.addTab(page, title_base)
            
        # Restore selection
        target_idx = -1
        if target_export_date:
            for i in range(self.history_tabs.count()):
                if self.history_tabs.tabText(i) == target_export_date:
                    target_idx = i
                    break
                    
        if target_idx == -1 and current_title_base:
            for i in range(self.history_tabs.count()):
                if self.history_tabs.tabText(i) == current_title_base:
                    target_idx = i
                    break
        
        if target_idx == -1 and self.history_tabs.count() > 0:
            # Fallback to NEW if no specific match
            target_idx = self.history_tabs.count() - 1
            
        if target_idx >= 0:
            self.history_tabs.setCurrentIndex(target_idx)

        self.history_tabs.blockSignals(False)

    def _build_history_card(self, item: HistoryItem) -> QFrame:
        card = QFrame()
        card.setObjectName("history-card")
        card.setStyleSheet("""
            QFrame { background: #fff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 14px; }
        """)

        layout = QHBoxLayout(card)
        layout.setSpacing(16)

        # Thumbnails
        thumb_row = QVBoxLayout()
        thumb_row.setSpacing(8)
        p_thumb = self._thumb_label(item.passport_path, "Passport")
        c_thumb = self._thumb_label(item.card_path, "Card")
        thumb_row.addWidget(p_thumb)
        thumb_row.addWidget(c_thumb)
        layout.addLayout(thumb_row)

        # Data
        data_layout = QVBoxLayout()
        data_layout.setSpacing(4)
        date_short = item.ts_iso[:10] if len(item.ts_iso) >= 10 else item.ts_iso
        date_lbl = QLabel(date_short)
        date_lbl.setStyleSheet("color: #64748b; font-size: 11px;")
        data_layout.addWidget(date_lbl)

        name = (item.summary or {}).get("Name") or "—"
        name_lbl = QLabel(str(name))
        name_lbl.setStyleSheet("font-weight: 600; color: #0f172a; font-size: 14px;")
        data_layout.addWidget(name_lbl)

        passport_no = (item.summary or {}).get("Passport Number") or "—"
        card_no = (item.summary or {}).get("Card Number") or "—"
        for txt in [f"Passport: {passport_no}", f"Card: {card_no}"]:
            sub = QLabel(txt)
            sub.setStyleSheet("color: #64748b; font-size: 12px;")
            data_layout.addWidget(sub)

        layout.addLayout(data_layout, 1)

        # Export button
        if item.combined:
            btn = QPushButton("Export to Excel")
            btn.setStyleSheet("background: #3b82f6; color: white;")
            btn.clicked.connect(lambda checked=False, h=item: self._export_history_item(h))
            layout.addWidget(btn, 0, Qt.AlignBottom)

        return card

    def _thumb_label(self, path: str, fallback: str) -> QLabel:
        lbl = QLabel()
        lbl.setFixedSize(80, 60)
        lbl.setAlignment(Qt.AlignCenter)
        lbl.setStyleSheet("background: #f1f5f9; border-radius: 8px; color: #94a3b8; font-size: 10px;")
        if path and Path(path).exists():
            pm = QPixmap(path)
            if not pm.isNull():
                scaled = pm.scaled(80, 60, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                lbl.setPixmap(scaled)
            else:
                lbl.setText(fallback)
        else:
            lbl.setText(fallback)
        return lbl

    def _export_history_item(self, item: HistoryItem) -> None:
        if not item.combined:
            QMessageBox.warning(self, "Export", "No data to export for this entry.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Save to Excel", "PASSPORT-FORM.xlsx", "Excel (*.xlsx)"
        )
        if not path:
            return
        try:
            from passport_data_extractor import PassportDataExtractor
            extractor = PassportDataExtractor(default_country_codes_path(), gpu=True)
            validity = self.validity_combo.currentText()
            extractor.save_to_excel(
                item.combined, path,
                validity_period=validity,
            )
            QMessageBox.information(self, "Export complete", f"Saved: {path}")
        except Exception as e:
            QMessageBox.critical(self, "Export failed", str(e))

    def _pick_passport(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Select passport image", "",
            "Images (*.png *.jpg *.jpeg *.webp *.bmp);;All files (*)",
        )
        if not path:
            return
        self._passport_path = path
        self.passport_preview.set_image(path)
        if hasattr(self, "btn_save_excel"):
            self.btn_save_excel.setEnabled(False)

    def _pick_card(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Select employment card image", "",
            "Images (*.png *.jpg *.jpeg *.webp *.bmp);;All files (*)",
        )
        if not path:
            return
        self._card_path = path
        self.card_preview.set_image(path)
        if hasattr(self, "btn_save_excel"):
            self.btn_save_excel.setEnabled(False)

    def _start_scan(self) -> None:
        if not self._passport_path and not self._card_path:
            QMessageBox.warning(
                self, "Missing file",
                "Please upload at least one image (passport or employee card).",
            )
            return
        try:
            self.progress.setValue(0)
            self.progress.setVisible(True)
            self.status_label.setText("Verifying…")
            self.btn_verify.setEnabled(False)
            self.btn_save_excel.setEnabled(False)

            # Clear fields before starting a new scan to prevent stale data display
            for w in list(self.passport_out.values()) + list(self.card_out.values()) + list(self.other_out.values()):
                w.clear()
            
            # Get potentially rotated paths from the previews
            passport = self.passport_preview.get_current_path() or ""
            card = self.card_preview.get_current_path() or ""
            
            self._start_worker(passport, card, "both")
        except Exception as e:  # noqa: BLE001
            # If anything goes wrong during setup, make sure the UI is not left stuck
            self.progress.setVisible(False)
            self.status_label.setText("Verification failed.")
            self.btn_verify.setEnabled(True)
            self.btn_save_excel.setEnabled(False)
            QMessageBox.critical(self, "Verification failed", str(e))


    def _clear(self) -> None:
        self._passport_path = None
        self._card_path = None
        self._last_result = None
        self.passport_preview.clear()
        self.card_preview.clear()
        self.progress.setValue(0)
        self.progress.setVisible(False)
        self.status_label.setText("Upload at least one image (passport and/or employee card), then click Verify.")
        for w in list(self.passport_out.values()) + list(self.card_out.values()) + list(self.other_out.values()):
            w.clear()
        self.validity_combo.setCurrentIndex(0)
        self.btn_save_excel.setEnabled(False)
        self.btn_verify.setEnabled(True)

    def _start_worker(self, passport_path: str, card_path: str, ocr_engine: str) -> None:
        if self._thread and self._thread.isRunning():
            return
        self._thread = QThread(self)
        self._worker = ExtractionWorker(
            country_codes_file=default_country_codes_path(),
            passport_path=passport_path,
            card_path=card_path,
            ocr_engine=ocr_engine,
            gpu=True,
            extractor=self._extractor,
        )

        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.progress.connect(self.progress.setValue)
        self._worker.status.connect(self.status_label.setText)
        self._worker.finished.connect(self._on_scan_finished)
        self._worker.failed.connect(self._on_scan_failed)
        self._worker.finished.connect(self._thread.quit)
        self._worker.failed.connect(self._thread.quit)
        self._thread.finished.connect(self._thread.deleteLater)
        self._worker.extractor_ready.connect(self._save_extractor)
        self._thread.start()

    def _save_extractor(self, extractor: object) -> None:
        self._extractor = extractor

    def _set_field_values(self, mapping: dict, data: dict, extra: dict | None = None) -> None:
        merged = {**(extra or {}), **data}
        for key, widget in mapping.items():
            v = merged.get(key, "Not Found")
            widget.setText("" if v in (None, "", "Not Found") else str(v))

    def _on_scan_finished(self, result_obj: object) -> None:
        result: ScanResult = result_obj
        self._last_result = result
        self._thread = None
        self._worker = None
        import datetime
        today = datetime.date.today()
        extra_dates = {
            'D01': f"{today.day:02d}",
            'D02': f"{today.month:02d}",
            'D03': f"{today.year}",
        }
        
        # Use combined data (Excel column names)
        self._set_field_values(self.passport_out, result.combined)
        self._set_field_values(self.card_out, result.combined)
        self._set_field_values(self.other_out, result.combined, extra=extra_dates)

        self.btn_save_excel.setEnabled(True)
        self.btn_verify.setEnabled(True)
        self.progress.setVisible(False)
        self.status_label.setText("Verification complete. Review and click 'Save Data to List'.")

    def _on_scan_failed(self, message: str) -> None:
        self._thread = None
        self._worker = None
        self.btn_verify.setEnabled(True)
        self.btn_save_excel.setEnabled(False)
        self.progress.setVisible(False)
        self.status_label.setText("Verification failed.")
        QMessageBox.critical(self, "Verification failed", message)

    def _recalculate_name02_from_fields(self) -> None:
        surname_text = self.passport_out.get("SURNAME").text().strip() if self.passport_out.get("SURNAME") else ""
        given_text = self.passport_out.get("GSURNAME").text().strip() if self.passport_out.get("GSURNAME") else ""
        name02 = " ".join([p for p in [surname_text, given_text] if p]).strip()

        self.passport_out["NAME 02"].setText(name02)

        if self._last_result:
            passport_data = dict(self._last_result.passport_data or {})
            passport_data["Surname"] = surname_text or passport_data.get("Surname", "")
            passport_data["Given Names"] = given_text or passport_data.get("Given Names", "")
            passport_data["Name"] = name02 or passport_data.get("Name", "")

            combined = dict(self._last_result.combined or {})
            combined["SURNAME"] = surname_text
            combined["GSURNAME"] = given_text
            combined["NAME 02"] = name02
            self._last_result = ScanResult(
                passport_data=passport_data,
                card_data=self._last_result.card_data,
                combined=combined,
            )

        self.status_label.setText("NAME 02 recalculated from SURNAME and GSURNAME.")
        self.btn_save_excel.setEnabled(True)

    def _save_data(self) -> None:
        result = self._last_result
        if not result:
            return
            
        current_data = dict(result.combined or {})
        for k, v in self.passport_out.items():
            current_data[k] = v.text().strip()
        for k, v in self.card_out.items():
            current_data[k] = v.text().strip()
        for k, v in self.other_out.items():
            current_data[k] = v.text().strip()
            
        import datetime
        today = datetime.date.today()
        # Only overwrite date if it wasn't manually typed in by user in the card_out
        if not current_data.get("D01"): current_data["D01"] = str(today.day).zfill(2)
        if not current_data.get("D02"): current_data["D02"] = str(today.month).zfill(2)
        if not current_data.get("D03"): current_data["D03"] = str(today.year)
        
        gender = current_data.get("Gender", "").upper().strip()
        current_data["F"] = ""
        current_data["M"] = ""
        if gender in ('F', 'FEMALE') or gender.startswith('F'):
            current_data["F"] = "X"
        elif gender in ('M', 'MALE') or gender.startswith('M'):
            current_data["M"] = "X"
            
        validity = self.validity_combo.currentText()
        current_data["12M"] = ""
        current_data["6M"] = ""
        current_data["3M"] = ""
        current_data["1M"] = ""
        if validity in ["12M", "6M", "3M", "1M"]:
            current_data[validity] = "X"
            
        summary = {
            "Name": current_data.get("NAME 02") or result.passport_data.get("Name"),
            "Passport Number": current_data.get("PASSPORT"),
            "Nationality": current_data.get("NASTIONALTY"),
            "Card Number": current_data.get("CARD NUMBER"),
        }
        
        self._history.append(
            passport_path=self._passport_path or "",
            card_path=self._card_path or "",
            ocr_engine="both",
            summary=summary,
            combined=current_data,
        )
        
        self.btn_save_excel.setEnabled(False)
        self.status_label.setText("Data saved to list. You can proceed with the next document.")
        QMessageBox.information(self, "Saved", "Data saved to list.")

    def _export_current_tab(self) -> None:
        items = self._history.load()
        if not items:
            QMessageBox.warning(self, "Export", "No data to export.")
            return

        idx = self.history_tabs.currentIndex()
        if idx < 0:
            return
            
        tab_text = self.history_tabs.tabText(idx)
        is_new = "New Data" in tab_text
        
        # Extract date from title like "2024-03-24 10:00:00 (5)"
        import re
        match = re.search(r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})', tab_text)
        target_date = match.group(1) if match else None

        filtered_items = []
        for it in items:
            if is_new:
                if not getattr(it, "exported", False):
                    filtered_items.append(it)
            else:
                if getattr(it, "exported", False) and it.export_date == target_date:
                    filtered_items.append(it)

        if not filtered_items:
            QMessageBox.warning(self, "Export", "No data to export in the current tab.")
            return

        data_list = []
        is_exported_list = []
        # We process them in order they appear in history (newest first usually)
        for it in filtered_items:
            if it.combined:
                data_list.append(it.combined)
                is_exported_list.append(it.exported)
                
        if not data_list:
            QMessageBox.warning(self, "Export", "No complete records to export in the current tab.")
            return
            
        path, _ = QFileDialog.getSaveFileName(
            self, "Export Current Tab", "PASSPORT-FORM.xlsx", "Excel (*.xlsx)"
        )
        if not path:
            return
            
        try:
            from passport_data_extractor import PassportDataExtractor
            extractor = PassportDataExtractor(default_country_codes_path(), gpu=False)
            
            validity = self.validity_combo.currentText()
            extractor.save_many_to_excel(
                data_list, path, 
                validity_period=validity,
                is_exported_list=is_exported_list
            )
            
            # Capture the export date to automatically follow it in the UI
            exp_date = self._history.mark_items_exported(filtered_items)
            self._refresh_history(target_export_date=exp_date)
            
            QMessageBox.information(self, "Export complete", f"Saved {len(data_list)} records from the current tab to: {path}")
        except Exception as e:
            QMessageBox.critical(self, "Export failed", str(e))

    def _clear_history(self) -> None:
        idx = self.history_tabs.currentIndex()
        if idx < 0:
            return
            
        tab_text = self.history_tabs.tabText(idx)
        is_new = "New Data" in tab_text
        
        msg = f"Delete all data in '{tab_text}'?"
        if QMessageBox.question(self, "Clear Tab", msg) != QMessageBox.Yes:
            return
            
        if is_new:
            self._history.clear_batch(exported=False)
        else:
            # Extract date from title like "2024-03-24 10:00:00 (5)"
            import re
            match = re.search(r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})', tab_text)
            date_str = match.group(1) if match else None
            self._history.clear_batch(exported=True, export_date=date_str)
            
        self._refresh_history()
