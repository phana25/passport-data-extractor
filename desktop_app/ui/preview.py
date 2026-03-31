from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable

from PySide6.QtCore import Qt, QRectF
from PySide6.QtGui import QBrush, QColor, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (
    QDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)


@dataclass(frozen=True)
class OverlayBox:
    label: str
    rect: QRectF  # in image pixel coordinates
    color: QColor = field(default_factory=lambda: QColor(32, 160, 72))


class ImagePreview(QFrame):
    def __init__(self) -> None:
        super().__init__()
        self.setFrameShape(QFrame.NoFrame)
        self.setMinimumHeight(80)
        self.setStyleSheet("background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px;")

        self._pixmap: QPixmap | None = None
        self._path: str | None = None
        self._boxes: list[OverlayBox] = []

    def set_image(self, path: str) -> None:
        pm = QPixmap(path)
        self._path = path
        self._pixmap = pm if not pm.isNull() else None
        self._boxes = []
        self.setCursor(Qt.PointingHandCursor if self._pixmap else Qt.ArrowCursor)
        self.update()

    def clear(self) -> None:
        self._path = None
        self._pixmap = None
        self._boxes = []
        self.setCursor(Qt.ArrowCursor)
        self.update()

    def set_overlay_boxes(self, boxes: list[OverlayBox]) -> None:
        self._boxes = boxes
        self.update()

    def paintEvent(self, event) -> None:  # noqa: N802
        super().paintEvent(event)
        p = QPainter(self)
        try:
            p.setRenderHint(QPainter.Antialiasing, True)

            if not self._pixmap:
                p.setPen(QColor(148, 163, 184))
                p.drawText(self.rect(), Qt.AlignCenter, "No image — click Browse above")
                return

            target = self._fit_rect(self._pixmap.width(), self._pixmap.height())
            # Use signature: drawPixmap(targetRect, pixmap, sourceRect)
            p.drawPixmap(target, self._pixmap, QRectF(self._pixmap.rect()))

            if not self._boxes:
                return

            # Map image pixel coords -> target rect coords
            sx = target.width() / self._pixmap.width()
            sy = target.height() / self._pixmap.height()
            ox = target.left()
            oy = target.top()

            for box in self._boxes:
                r = QRectF(
                    ox + box.rect.left() * sx,
                    oy + box.rect.top() * sy,
                    box.rect.width() * sx,
                    box.rect.height() * sy,
                )

                pen = QPen(box.color)
                pen.setWidth(2)
                p.setPen(pen)
                p.setBrush(Qt.NoBrush)
                p.drawRoundedRect(r, 6, 6)

                tag_rect = QRectF(r.left(), max(target.top(), r.top() - 22), min(170, r.width()), 18)
                p.setPen(Qt.NoPen)
                p.setBrush(QBrush(box.color))
                p.drawRoundedRect(tag_rect, 6, 6)
                p.setPen(QColor("white"))
                p.drawText(tag_rect, Qt.AlignCenter, box.label)
        finally:
            p.end()

    def _fit_rect(self, img_w: int, img_h: int):
        # Scale to fill width of box, use full width
        w = max(1, self.width() - 8)
        h = max(1, self.height() - 8)
        if img_w <= 0 or img_h <= 0:
            return self.rect()
        scale = w / img_w  # prioritize width to fill box
        scale = min(scale, h / img_h)  # cap by height so it fits
        tw = int(img_w * scale)
        th = int(img_h * scale)
        x = (self.width() - tw) // 2
        y = (self.height() - th) // 2
        return QRectF(x, y, tw, th)

    def mousePressEvent(self, event):
        if self._pixmap and event.button() == Qt.LeftButton:
            self._on_click()
        super().mousePressEvent(event)

    def _on_click(self) -> None:
        """Override in subclass or connect to show full-size preview."""
        pass


class FullSizePreviewDialog(QDialog):
    """Modal dialog showing image at full/near-full size with scroll."""

    def __init__(self, path: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Image Preview")
        self.setModal(True)
        self.setMinimumSize(400, 300)
        self.resize(800, 600)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("background: #1e293b;")

        lbl = QLabel()
        lbl.setAlignment(Qt.AlignCenter)
        lbl.setStyleSheet("background: #1e293b;")
        pm = QPixmap(path)
        if not pm.isNull():
            # Show at up to 1600px wide, keep aspect
            max_w = 1600
            if pm.width() > max_w:
                scaled = pm.scaledToWidth(max_w, Qt.SmoothTransformation)
                lbl.setPixmap(scaled)
            else:
                lbl.setPixmap(pm)
        else:
            lbl.setText("Could not load image")
            lbl.setStyleSheet("background: #1e293b; color: #94a3b8;")

        scroll.setWidget(lbl)
        layout.addWidget(scroll)

        btn_row = QWidget()
        btn_layout = QVBoxLayout(btn_row)
        btn_layout.setContentsMargins(16, 12, 16, 16)
        close_btn = QPushButton("Close")
        close_btn.setStyleSheet(
            "QPushButton { background: #475569; color: white; padding: 8px 24px; border-radius: 8px; } "
            "QPushButton:hover { background: #64748b; }"
        )
        close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(close_btn, 0, Qt.AlignCenter)
        layout.addWidget(btn_row)


class UploadPreviewZone(QFrame):
    """Combined upload area and preview: shows placeholder when empty, image when loaded."""

    def __init__(
        self,
        title: str,
        on_browse: Callable[[], None],
        on_clear: Callable[[], None] | None = None,
    ) -> None:
        super().__init__()
        self.setObjectName("UploadPreviewZone")
        self.setFrameShape(QFrame.NoFrame)
        self.setMinimumHeight(120)
        self.setMaximumHeight(260)
        self._on_browse = on_browse
        self._on_clear = on_clear or (lambda: None)
        self._title = title

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)

        self._stack = QStackedWidget()

        # Empty state
        empty = QWidget()
        empty_layout = QVBoxLayout(empty)
        empty_layout.setAlignment(Qt.AlignCenter)
        empty_layout.setSpacing(12)
        lbl = QLabel(f"Click to add {title}")
        lbl.setStyleSheet("color: #94a3b8; font-size: 14px;")
        empty_layout.addWidget(lbl)
        btn = QPushButton("Browse…")
        btn.setStyleSheet(
            "QPushButton { background: #3b82f6; color: white; padding: 8px 20px; border-radius: 8px; } "
            "QPushButton:hover { background: #2563eb; }"
        )
        btn.clicked.connect(on_browse)
        empty_layout.addWidget(btn, 0, Qt.AlignCenter)
        self._stack.addWidget(empty)

        # Loaded state: ImagePreview + Cancel / Re-upload buttons
        loaded = QWidget()
        loaded_layout = QGridLayout(loaded)
        loaded_layout.setContentsMargins(0, 0, 0, 0)

        self._preview = ImagePreview()
        self._preview.setMinimumHeight(140)

        def show_full_preview():
            if self._preview._path:
                dlg = FullSizePreviewDialog(self._preview._path, self)
                dlg.exec()

        self._preview._on_click = show_full_preview
        self._preview.setStyleSheet("background: transparent; border: none;")
        loaded_layout.addWidget(self._preview, 0, 0)

        btn_col = QVBoxLayout()
        btn_col.setContentsMargins(12, 12, 12, 12)
        btn_col.setSpacing(6)
        
        reupload_btn = QPushButton("Re-upload")
        reupload_btn.setStyleSheet(
            "QPushButton { background: #3b82f6; color: white; padding: 6px 14px; border-radius: 6px; font-size: 12px; } "
            "QPushButton:hover { background: #2563eb; }"
        )
        reupload_btn.clicked.connect(on_browse)
        
        cancel_btn = QPushButton("Cancel")
        cancel_btn.setStyleSheet(
            "QPushButton { background: #e2e8f0; color: #475569; padding: 6px 14px; border-radius: 6px; font-size: 12px; } "
            "QPushButton:hover { background: #cbd5e1; }"
        )
        cancel_btn.clicked.connect(self._do_clear)
        
        btn_col.addWidget(reupload_btn)
        btn_col.addWidget(cancel_btn)
        
        loaded_layout.addLayout(btn_col, 0, 0, Qt.AlignLeft | Qt.AlignBottom)

        self._stack.addWidget(loaded)

        layout.addWidget(self._stack, 1)

    def set_image(self, path: str) -> None:
        self._preview.set_image(path)
        self._stack.setCurrentIndex(1)

    def clear(self) -> None:
        self._preview.clear()
        self._stack.setCurrentIndex(0)

    def _do_clear(self) -> None:
        self.clear()
        self._on_clear()

    def set_overlay_boxes(self, boxes: list) -> None:
        self._preview.set_overlay_boxes(boxes)

