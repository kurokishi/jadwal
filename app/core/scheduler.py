import pandas as pd
from datetime import datetime, timedelta
from .time_parser import TimeParser
from .cleaner import DataCleaner


class Scheduler:

    def __init__(self, config):
        self.config = config
        self.tp = TimeParser()

    def generate_slots(self):
        slots = []
        current = datetime.strptime(
            f"{self.config.start_hour}:{self.config.start_minute}", "%H:%M"
        )

        end_time = self.config.time_slot_end()

        while current.time() <= end_time:
            slots.append(current.time())
            current += timedelta(minutes=self.config.interval_minutes)

        return slots

    # FINAL API METHOD
    def process_schedule(self, df, jenis):
        df = DataCleaner.clean(df, self.config.hari_list, jenis, self.config.auto_fix_errors)

        if df.empty:
            return pd.DataFrame()

        slots = self.generate_slots()
        slot_str = [t.strftime("%H:%M") for t in slots]

        results = []

        for (dok, poli), group in df.groupby(["Nama Dokter", "Poli Asal"]):
            for hari in self.config.hari_list:
                if hari not in group.columns:
                    continue

                ranges = []
                for s in group[hari].dropna():
                    start, end = self.tp.parse(s)
                    if start and end:
                        ranges.append((start, end))

                if not ranges:
                    continue

                merged = self.merge_ranges(ranges)

                row = {
                    "POLI ASAL": poli,
                    "JENIS POLI": jenis,
                    "HARI": hari,
                    "DOKTER": dok
                }

                for i, sl in enumerate(slots):
                    sl_end = (datetime.combine(datetime.today(), sl) +
                              timedelta(minutes=self.config.interval_minutes)).time()

                    overlap = any(not (sl_end <= a or sl >= b) for a, b in merged)

                    row[slot_str[i]] = 'R' if overlap and jenis == "Reguler" else \
                                       'E' if overlap and jenis != "Reguler" else ""

                results.append(row)

        df_out = pd.DataFrame(results)
        return df_out

    @staticmethod
    def merge_ranges(ranges):
        if not ranges:
            return []

        ranges = sorted(ranges, key=lambda x: x[0])
        merged = [list(ranges[0])]

        for start, end in ranges[1:]:
            last_start, last_end = merged[-1]
            if start <= last_end:
                merged[-1][1] = max(last_end, end)
            else:
                merged.append([start, end])

        return merged
