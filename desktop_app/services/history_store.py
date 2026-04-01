from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

from PySide6.QtCore import QStandardPaths


@dataclass
class HistoryItem:
    ts_iso: str
    passport_path: str
    card_path: str
    ocr_engine: str
    summary: dict
    combined: dict | None = None  # Full extraction for re-export to Excel
    exported: bool = False
    export_date: str | None = None


class HistoryStore:
    def __init__(self, filename: str = "history.json") -> None:
        base = Path(QStandardPaths.writableLocation(QStandardPaths.AppDataLocation))
        base.mkdir(parents=True, exist_ok=True)
        self.path = base / filename

    def load(self) -> list[HistoryItem]:
        if not self.path.exists():
            return []
        try:
            raw = json.loads(self.path.read_text(encoding="utf-8"))
            out: list[HistoryItem] = []
            for it in raw if isinstance(raw, list) else []:
                if not isinstance(it, dict):
                    continue
                summary = it.get("summary", {})
                if not isinstance(summary, dict):
                    summary = {}
                combined = it.get("combined")
                if not isinstance(combined, dict):
                    combined = None
                out.append(
                    HistoryItem(
                        ts_iso=str(it.get("ts_iso", "")),
                        passport_path=str(it.get("passport_path", "")),
                        card_path=str(it.get("card_path", "")),
                        ocr_engine=str(it.get("ocr_engine", "")),
                        summary=summary,
                        combined=combined,
                        exported=bool(it.get("exported", False)),
                        export_date=it.get("export_date"),
                    )
                )
            return out
        except Exception:  # noqa: BLE001
            return []

    def append(
        self,
        passport_path: str,
        card_path: str,
        ocr_engine: str,
        summary: dict,
        combined: dict | None = None,
    ) -> None:
        items = self.load()
        now = datetime.now(timezone.utc).isoformat()
        items.insert(
            0,
            HistoryItem(
                ts_iso=now,
                passport_path=passport_path,
                card_path=card_path,
                ocr_engine=ocr_engine,
                summary=summary,
                combined=combined,
                exported=False,
            ),
        )
        # keep last 100
        items = items[:100]
        self._save(items)

    def _save(self, items: list[HistoryItem]) -> None:
        out = []
        for i in items:
            d = {
                "ts_iso": i.ts_iso,
                "passport_path": i.passport_path,
                "card_path": i.card_path,
                "ocr_engine": i.ocr_engine,
                "summary": i.summary,
                "exported": i.exported,
            }
            if getattr(i, "export_date", None):
                d["export_date"] = i.export_date
            if i.combined is not None:
                d["combined"] = i.combined
            out.append(d)
        self.path.write_text(
            json.dumps(out, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        
    def mark_items_exported(self, items_to_mark: list[HistoryItem]) -> str | None:
        """Mark a specific subset of items as exported with the current timestamp."""
        all_items = self.load()
        from datetime import datetime
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Match items by their timestamp ISO string
        target_ts = {it.ts_iso for it in items_to_mark}
        
        changed = False
        final_date = None
        for i in all_items:
            if i.ts_iso in target_ts:
                if not i.exported:
                    i.exported = True
                    i.export_date = now
                    changed = True
                final_date = i.export_date
        
        if changed:
            self._save(all_items)
        
        # If we just moved items from NEW, return 'now'. If we re-exported an old tab, return its date.
        return final_date if final_date else (now if changed else None)

    def mark_all_exported(self) -> None:
        items = self.load()
        self.mark_items_exported(items)

    def clear_batch(self, exported: bool, export_date: str | None = None) -> None:
        """Clear only items matching the exported status and optional date."""
        items = self.load()
        if not exported:
            # Clear all unexported ("NEW") items
            new_items = [i for i in items if i.exported]
        else:
            # Clear items matching this specific export date
            new_items = [i for i in items if not (i.exported and i.export_date == export_date)]
            
        self._save(new_items)

    def clear(self) -> None:
        try:
            if self.path.exists():
                self.path.unlink()
        except Exception:  # noqa: BLE001
            # Best-effort; UI will just show what's still on disk.
            pass

