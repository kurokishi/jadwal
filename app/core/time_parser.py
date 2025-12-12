import re
import pandas as pd
from datetime import time, datetime

class TimeParser:
    def __init__(self, start_hour=7, start_minute=30, interval_minutes=30):
        self.start_hour = start_hour
        self.start_minute = start_minute
        self.interval_minutes = interval_minutes
    
    @staticmethod
    def parse(time_str):
        """Static method untuk parse waktu (untuk backward compatibility)"""
        if pd.isna(time_str) or str(time_str).strip() == "":
            return None, None

        s = str(time_str).strip().replace(" ", "").replace(".", ":")

        m = re.search(r"(\d{1,2}:\d{2})-(\d{1,2}:\d{2})", s)
        if not m:
            return None, None

        try:
            sh, sm = map(int, m.group(1).split(":"))
            eh, em = map(int, m.group(2).split(":"))
            return time(sh, sm), time(eh, em)
        except:
            return None, None
    
    def parse_time_range(self, time_range_str, slot_strings):
        """
        Parse string waktu ke list slot berdasarkan konfigurasi
        """
        start_time, end_time = self.parse(time_range_str)
        if not start_time or not end_time:
            return []
        
        result_slots = []
        for slot in slot_strings:
            slot_time = datetime.strptime(slot, "%H:%M").time()
            if start_time <= slot_time < end_time:
                result_slots.append(slot)
        
        return result_slots
    
    def generate_slot_strings(self):
        """
        Generate list slot berdasarkan konfigurasi start time dan interval
        """
        slot_strings = []
        current_time = self.start_hour * 60 + self.start_minute
        
        # Generate sampai jam 14:30
        end_time = 14 * 60 + 30
        
        while current_time < end_time:
            hours = current_time // 60
            minutes = current_time % 60
            time_str = f"{hours:02d}:{minutes:02d}"
            slot_strings.append(time_str)
            current_time += self.interval_minutes
        
        return slot_strings
