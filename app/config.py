from dataclasses import dataclass, field
from datetime import time


@dataclass
class Config:
    start_hour: int = 7
    start_minute: int = 30
    interval_minutes: int = 30
    max_poleks_per_slot: int = 7
    auto_fix_errors: bool = True
    enable_sabtu: bool = False

    hari_order: dict = field(default_factory=lambda: {
        "Senin": 1,
        "Selasa": 2,
        "Rabu": 3,
        "Kamis": 4,
        "Jum'at": 5
    })

    @property
    def hari_list(self):
        hari = list(self.hari_order.keys())
        if self.enable_sabtu and "Sabtu" not in hari:
            hari.append("Sabtu")
        return hari

    def time_slot_end(self):
        return time(14, 30)
