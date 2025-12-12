# debug.py
import pandas as pd
import io
import sys
import os

# Tambahkan path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from app.config import Config
from app.core.scheduler import Scheduler
from app.core.cleaner import DataCleaner
from app.core.time_parser import TimeParser

# Test dengan file Anda
config = Config()
time_parser = TimeParser(
    start_hour=config.start_hour,
    start_minute=config.start_minute,
    interval_minutes=config.interval_minutes
)
cleaner = DataCleaner()
scheduler = Scheduler(time_parser, cleaner, config)

# Load file Anda
file_path = "View jadwal (1).xlsx"  # Ganti dengan path file Anda
with open(file_path, 'rb') as f:
    file_bytes = f.read()

print("Testing file processing...")
grid_df, slot_strings, errors = scheduler.process_dataframe(io.BytesIO(file_bytes))

print(f"\nResults:")
print(f"- Grid DF: {'Exists' if grid_df is not None else 'None'}")
print(f"- Grid shape: {grid_df.shape if grid_df is not None else 'N/A'}")
print(f"- Slots: {len(slot_strings) if slot_strings else 0}")
print(f"- Errors: {len(errors)}")

if errors:
    print("\nErrors:")
    for error in errors:
        print(f"- {error}")

if grid_df is not None:
    print(f"\nFirst few rows:")
    print(grid_df.head())
    print(f"\nColumns: {list(grid_df.columns)}")
