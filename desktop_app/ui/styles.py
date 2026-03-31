"""Modern app styles."""
from __future__ import annotations

STYLESHEET = """
/* Root */
QMainWindow, QWidget#root {
    background-color: #f5f6f8;
}
QLabel {
    background: transparent;
}

/* Nav bar */
QFrame#nav-bar {
    background-color: #ffffff;
    border: none;
    border-bottom: 2px solid #f1f5f9;
}
QLabel#sidebar-title {
    color: #1e293b;
    font-size: 16px;
    font-weight: 700;
}
QPushButton#nav-btn {
    background: transparent;
    color: #475569;
    padding: 0 16px;
    border-radius: 8px;
    font-size: 14px;
    font-weight: 500;
}
QPushButton#nav-btn:hover {
    background: #f8fafc;
    color: #3b82f6;
}
QPushButton#nav-btn:checked {
    background: #eff6ff;
    color: #2563eb;
    font-weight: 600;
}
QLabel#theme-label {
    color: #64748b;
    font-size: 13px;
    font-weight: 500;
}
QComboBox {
    combobox-popup: 0;
}
QComboBox#theme-combo {
    background: #f8fafc;
    color: #1e293b;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    padding: 4px 12px;
    font-size: 13px;
}
QComboBox#theme-combo::drop-down {
    subcontrol-origin: padding;
    border: none;
    width: 24px;
}
QComboBox#theme-combo::down-arrow {
    image: none;
    border: none;
}

/* Content */
QLabel#page-title {
    color: #0f172a;
    font-size: 20px;
    font-weight: 700;
}

/* Cards / Panels */
QFrame#card, QGroupBox {
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 16px;
    margin: 4px 0;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 14px;
    padding: 0 8px;
    background: #ffffff;
    color: #475569;
    font-weight: 600;
    font-size: 11px;
}

/* Buttons */
QPushButton {
    background-color: #3b82f6;
    color: white;
    border: none;
    border-radius: 8px;
    padding: 10px 18px;
    font-size: 13px;
    font-weight: 500;
}
QPushButton:hover {
    background-color: #2563eb;
}
QPushButton:pressed {
    background-color: #1d4ed8;
}
QPushButton:disabled {
    background-color: #94a3b8;
    color: #cbd5e1;
}
QPushButton#secondary {
    background-color: #e2e8f0;
    color: #475569;
}
QPushButton#secondary:hover {
    background-color: #cbd5e1;
}
QPushButton#secondary:pressed {
    background-color: #94a3b8;
}
QPushButton#danger {
    background-color: #ef4444;
}
QPushButton#danger:hover {
    background-color: #dc2626;
}

/* Inputs */
QLineEdit {
    background-color: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    padding: 8px 12px;
    font-size: 13px;
    color: #0f172a;
}
QLineEdit:focus {
    border-color: #3b82f6;
}
QComboBox {
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    padding: 8px 12px;
    min-width: 120px;
    color: #0f172a;
}
QComboBox QAbstractItemView {
    background-color: #ffffff;
    color: #0f172a;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    outline: none;
    padding: 4px;
}
QComboBox QAbstractItemView::item {
    background-color: transparent;
    padding: 8px 12px;
    border-radius: 4px;
    color: #0f172a;
}
QComboBox QAbstractItemView::item:selected,
QComboBox QAbstractItemView::item:hover {
    background-color: #3b82f6;
    color: #ffffff;
}

/* Progress */
QProgressBar {
    border: none;
    background-color: #e2e8f0;
    border-radius: 6px;
    height: 8px;
    text-align: center;
}
QProgressBar::chunk {
    background-color: #3b82f6;
    border-radius: 6px;
}

/* Upload zone */
QFrame#UploadPreviewZone {
    background: #f8fafc;
    border: 2px dashed #cbd5e1;
    border-radius: 12px;
}
QFrame#UploadPreviewZone:hover {
    border-color: #94a3b8;
}

/* History card */
QFrame#history-card {
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 14px;
}
QLabel#history-date {
    color: #64748b;
    font-size: 12px;
}
QLabel#history-name {
    color: #0f172a;
    font-size: 14px;
    font-weight: 600;
}

/* Data section (light) */
QFrame#data-section-box {
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
}
QLineEdit#data-field {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    padding: 8px 12px;
    font-size: 13px;
    color: #0f172a;
}
QLineEdit#data-field:focus {
    border-color: #3b82f6;
    background: #ffffff;
}
QFrame#opts-bar {
}
QComboBox#validity-combo {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 13px;
}
QComboBox#validity-combo QAbstractItemView {
    background: #ffffff;
    color: #0f172a;
}
QLabel#data-section-header {
    color: #94a3b8;
    font-size: 11px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-bottom: 8px;
}
QLabel#data-field-label {
    color: #64748b;
    font-size: 10px;
    font-weight: 700;
}
QLabel#opts-label {
    color: #64748b;
    font-size: 13px;
}
"""

STYLESHEET_DARK = """
/* Root */
QMainWindow, QWidget#root {
    background-color: #0f172a;
}
QLabel {
    background: transparent;
}

/* Nav bar */
QFrame#nav-bar {
    background-color: #0f172a;
    border-bottom: 2px solid #1e293b;
}
QLabel#sidebar-title {
    color: #f8fafc;
    font-size: 16px;
    font-weight: 700;
}
QPushButton#nav-btn {
    background: transparent;
    color: #94a3b8;
    padding: 0 16px;
    border-radius: 8px;
    font-size: 14px;
    font-weight: 500;
}
QPushButton#nav-btn:hover {
    background: #1e293b;
    color: #f8fafc;
}
QPushButton#nav-btn:checked {
    background: #3b82f6;
    color: #ffffff;
    font-weight: 600;
}
QLabel#theme-label {
    color: #94a3b8;
    font-size: 13px;
    font-weight: 500;
}
QComboBox {
    combobox-popup: 0;
}
QComboBox#theme-combo {
    background: #1e293b;
    color: #f8fafc;
    border: 1px solid #334155;
    border-radius: 8px;
    padding: 4px 12px;
    font-size: 13px;
}
QComboBox#theme-combo::drop-down {
    subcontrol-origin: padding;
    border: none;
    width: 24px;
}
QComboBox#theme-combo::down-arrow {
    image: none;
    border: none;
}

/* Content */
QLabel#page-title {
    color: #f8fafc;
    font-size: 20px;
    font-weight: 700;
}

/* Cards / Panels */
QFrame#card, QGroupBox {
    background-color: #1e293b;
    border: 1px solid #334155;
    border-radius: 12px;
    padding: 16px;
    margin: 4px 0;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 14px;
    padding: 0 8px;
    background: #1e293b;
    color: #94a3b8;
    font-weight: 600;
    font-size: 11px;
}

/* Buttons */
QPushButton {
    background-color: #3b82f6;
    color: white;
    border: none;
    border-radius: 8px;
    padding: 10px 18px;
    font-size: 13px;
    font-weight: 500;
}
QPushButton:hover {
    background-color: #2563eb;
}
QPushButton:pressed {
    background-color: #1d4ed8;
}
QPushButton:disabled {
    background-color: #475569;
    color: #94a3b8;
}
QPushButton#secondary {
    background-color: #334155;
    color: #e2e8f0;
}
QPushButton#secondary:hover {
    background-color: #475569;
}
QPushButton#secondary:pressed {
    background-color: #64748b;
}
QPushButton#danger {
    background-color: #dc2626;
}
QPushButton#danger:hover {
    background-color: #b91c1c;
}

/* Inputs */
QLineEdit {
    background-color: #1e293b;
    border: 1px solid #334155;
    border-radius: 8px;
    padding: 8px 12px;
    font-size: 13px;
    color: #f8fafc;
}
QLineEdit:focus {
    border-color: #3b82f6;
}
QComboBox {
    background-color: #1e293b;
    border: 1px solid #334155;
    border-radius: 8px;
    padding: 8px 12px;
    min-width: 120px;
    color: #f8fafc;
}
QComboBox QAbstractItemView {
    background-color: #1e293b;
    color: #f8fafc;
    border: 1px solid #334155;
    border-radius: 8px;
    outline: none;
    padding: 4px;
}
QComboBox QAbstractItemView::item {
    background-color: transparent;
    padding: 8px 12px;
    border-radius: 4px;
    color: #f8fafc;
}
QComboBox QAbstractItemView::item:selected,
QComboBox QAbstractItemView::item:hover {
    background-color: #3b82f6;
    color: #ffffff;
}

/* Upload zone (dark) */
QFrame#UploadPreviewZone {
    background: #1e293b;
    border: 2px dashed #475569;
    border-radius: 12px;
}
QFrame#UploadPreviewZone:hover {
    border-color: #64748b;
}

/* Progress */
QProgressBar {
    border: none;
    background-color: #334155;
    border-radius: 6px;
    height: 8px;
    text-align: center;
}
QProgressBar::chunk {
    background-color: #3b82f6;
    border-radius: 6px;
}

/* History card */
QFrame#history-card {
    background-color: #1e293b;
    border: 1px solid #334155;
    border-radius: 12px;
    padding: 14px;
}
QLabel#history-date {
    color: #94a3b8;
    font-size: 12px;
}
QLabel#history-name {
    color: #f8fafc;
    font-size: 14px;
    font-weight: 600;
}

/* Data section (dark) */
QFrame#data-section-box {
    background: #1e293b;
    border: 1px solid #334155;
    border-radius: 12px;
}
QLineEdit#data-field {
    background: #0f172a;
    border: 1px solid #334155;
    border-radius: 8px;
    padding: 8px 12px;
    font-size: 13px;
    color: #f8fafc;
}
QLineEdit#data-field:focus {
    border-color: #3b82f6;
    background: #0f172a;
}
QFrame#opts-bar {
}
QComboBox#validity-combo {
    background: #0f172a;
    border: 1px solid #334155;
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 13px;
    color: #f8fafc;
}
QComboBox#validity-combo QAbstractItemView {
    background: #1e293b;
    color: #f8fafc;
}
QLabel#data-section-header {
    color: #94a3b8;
    font-size: 11px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-bottom: 8px;
}
QLabel#data-field-label {
    color: #94a3b8;
    font-size: 10px;
    font-weight: 700;
}
QLabel#opts-label {
    color: #94a3b8;
    font-size: 13px;
}
"""


def get_stylesheet(theme: str) -> str:
    """Return stylesheet for theme: 'light', 'dark', or 'system'."""
    if theme == "dark":
        return STYLESHEET_DARK
    if theme == "system":
        try:
            from PySide6.QtCore import Qt
            from PySide6.QtGui import QGuiApplication
            hints = QGuiApplication.styleHints()
            if hasattr(hints, "colorScheme"):
                cs = hints.colorScheme()
                if cs == Qt.ColorScheme.Dark:
                    return STYLESHEET_DARK
        except Exception:
            pass
    return STYLESHEET
