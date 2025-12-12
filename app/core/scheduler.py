# app/core/scheduler.py - Tambahkan method untuk handle bytes
import pandas as pd
import io
from typing import Dict, List, Tuple

class Scheduler:
    def __init__(self, parser, cleaner, config):
        self.parser = parser
        self.cleaner = cleaner
        self.config = config
        
    def process_dataframe(self, df_or_file):
        """
        Proses data dari file Excel atau DataFrame
        """
        error_messages = []
        
        try:
            # 1. Clean data (method cleaner sudah handle file Excel)
            cleaned_df = self.cleaner.clean(df_or_file)
            
            if cleaned_df.empty:
                error_messages.append("Data setelah cleaning kosong")
                return None, [], error_messages
            
            # 2. Generate slot strings
            slot_strings = self._generate_slot_strings()
            
            # 3. Parse waktu ke slot
            slot_df = self._parse_time_to_slots(cleaned_df, slot_strings)
            
            # 4. Create grid format
            grid_df = self._create_grid_format(slot_df, slot_strings)
            
            # 5. Validasi
            validation_errors = self._validate_grid(grid_df, slot_strings)
            if validation_errors:
                error_messages.extend(validation_errors)
            
            return grid_df, slot_strings, error_messages
            
        except Exception as e:
            error_messages.append(f"Error processing data: {str(e)}")
            return None, [], error_messages
    
    # ... method lainnya tetap sama ...
