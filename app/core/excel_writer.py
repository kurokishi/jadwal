"""
ExcelWriter - Modul untuk menulis hasil jadwal ke file Excel
Membuat file Excel dengan multiple sheets: Jadwal, Rekap, Analisis, dll.
"""

import io
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
import traceback


class ExcelWriter:
    def __init__(self, config):
        """
        Inisialisasi ExcelWriter dengan konfigurasi
        
        Args:
            config: Config object berisi pengaturan aplikasi
        """
        self.config = config
        self.interval = config.interval_minutes
        self.max_e = config.max_poleks_per_slot
        
        print(f"✅ ExcelWriter initialized with config:")
        print(f"   - Interval: {self.interval} minutes")
        print(f"   - Max poleks per slot: {self.max_e}")
        
        # ======================================================
        # DEFINE COLORS
        # ======================================================
        
        # Warna slot
        self.fill_r = PatternFill("solid", fgColor="C6EFCE")   # Hijau soft - Reguler
        self.fill_e = PatternFill("solid", fgColor="BDD7EE")   # Biru soft - Poleks
        self.fill_over = PatternFill("solid", fgColor="FFC7CE")  # Merah soft - Overload
        
        # Warna konflik
        self.fill_conflict = PatternFill("solid", fgColor="FFD966")  # Kuning - Konflik ringan
        self.fill_conflict_hard = PatternFill("solid", fgColor="FF0000")  # Merah - Konflik berat
        
        # Warna lainnya
        self.fill_header = PatternFill("solid", fgColor="366092")  # Biru tua - Header
        self.fill_gray = PatternFill("solid", fgColor="F2F2F2")    # Abu-abu - Alternating rows
        self.fill_total = PatternFill("solid", fgColor="FFE699")   # Kuning muda - Total
        
        # Font
        self.font_header = Font(bold=True, color="FFFFFF", size=11)
        self.font_normal = Font(size=10)
        self.font_bold = Font(bold=True)
        self.font_small = Font(size=9)
        
        # Alignment
        self.align_center = Alignment(horizontal="center", vertical="center")
        self.align_left = Alignment(horizontal="left", vertical="center")
        self.align_right = Alignment(horizontal="right", vertical="center")
        
        # Borders
        self.thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )
        
        self.thick_border = Border(
            bottom=Side(style="thick")
        )
        
        print("✅ ExcelWriter styling configured")
    
    # ======================================================
    # MAIN WRITE METHOD
    # ======================================================
    
    def write(self, source_file, df_grid, slot_str):
        """
        Tulis hasil jadwal ke file Excel dengan multiple sheets
        
        Args:
            source_file: File Excel asli (BytesIO atau path) sebagai template
            df_grid: DataFrame hasil scheduler (format grid)
            slot_str: List string slot waktu
            
        Returns:
            BytesIO buffer berisi file Excel
        """
        print(f"📝 ExcelWriter.write() called")
        print(f"   - df_grid shape: {df_grid.shape if df_grid is not None else 'None'}")
        print(f"   - slot_str length: {len(slot_str) if slot_str else 0}")
        
        try:
            # Load workbook dari source file
            wb = self._load_workbook(source_file)
            
            if wb is None:
                print("❌ Failed to load workbook, creating new one")
                wb = Workbook()
            
            print(f"✅ Workbook loaded/created with {len(wb.sheetnames)} sheets")
            
            # Buat semua sheets
            print("1. Creating 'Jadwal' sheet...")
            self._create_jadwal_sheet(wb, df_grid, slot_str)
            
            print("2. Creating 'Rekap Layanan' sheet...")
            self._create_rekap_layanan_sheet(wb, df_grid, slot_str)
            
            print("3. Creating 'Rekap Poli' sheet...")
            self._create_rekap_poli_sheet(wb, df_grid, slot_str)
            
            print("4. Creating 'Rekap Dokter' sheet...")
            self._create_rekap_dokter_sheet(wb, df_grid, slot_str)
            
            print("5. Creating 'Peak Hour Analysis' sheet...")
            self._create_peak_hour_sheet(wb, df_grid, slot_str)
            
            print("6. Creating 'Conflict Dokter' sheet...")
            self._create_conflict_doctor_sheet(wb, df_grid, slot_str)
            
            print("7. Creating 'Peta Konflik Dokter' sheet...")
            self._create_conflict_map_sheet(wb, df_grid, slot_str)
            
            print("8. Creating 'Grafik Poli' sheet...")
            self._create_grafik_poli_sheet(wb, df_grid)
            
            print("9. Creating 'Summary' sheet...")
            self._create_summary_sheet(wb, df_grid, slot_str)
            
            # Apply styling ke semua sheets
            print("10. Applying styling to all sheets...")
            self._apply_styling_to_all_sheets(wb)
            
            # Auto adjust column widths
            print("11. Auto-adjusting column widths...")
            self._auto_adjust_column_widths(wb)
            
            # Save to buffer
            print("12. Saving to buffer...")
            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            
            print(f"✅ Excel file created successfully: {buf.getbuffer().nbytes} bytes")
            return buf
            
        except Exception as e:
            print(f"❌ Error in ExcelWriter.write(): {e}")
            print(traceback.format_exc())
            raise
    
    # ======================================================
    # SHEET CREATION METHODS
    # ======================================================
    
    def _create_jadwal_sheet(self, wb, df_grid, slot_str):
        """Buat sheet Jadwal utama"""
        # Hapus sheet lama jika ada
        if "Jadwal" in wb.sheetnames:
            ws_index = wb.sheetnames.index("Jadwal")
            del wb[wb.sheetnames[ws_index]]
        
        ws = wb.create_sheet("Jadwal", 0)  # Buat sebagai sheet pertama
        
        # Header
        headers = ["POLI", "JENIS", "HARI", "DOKTER", "JAM"] + slot_str
        ws.append(headers)
        
        # Data
        if df_grid is not None and not df_grid.empty:
            for _, row in df_grid.iterrows():
                ws.append([row.get(h, "") for h in headers])
        
        # Apply styling
        self._style_jadwal_sheet(ws, df_grid, slot_str)
        
        # Freeze header row
        ws.freeze_panes = "F2"
        
        return ws
    
    def _create_rekap_layanan_sheet(self, wb, df_grid, slot_str):
        """Buat sheet Rekap Layanan"""
        if "Rekap Layanan" in wb.sheetnames:
            del wb["Rekap Layanan"]
        
        ws = wb.create_sheet("Rekap Layanan")
        ws.append(["POLI", "HARI", "DOKTER", "JENIS", "WAKTU LAYANAN"])
        
        if df_grid is not None and not df_grid.empty:
            for (poli, hari, dokter, jenis), group in df_grid.groupby(["POLI", "HARI", "DOKTER", "JENIS"]):
                active_slots = [s for s in slot_str if s in group.columns and group.iloc[0].get(s) in ["R", "E"]]
                
                if active_slots:
                    time_ranges = self._combine_slots_to_ranges(active_slots)
                    
                    for time_range in time_ranges:
                        ws.append([poli, hari, dokter, jenis, time_range])
        
        # Style
        self._style_rekap_sheet(ws)
        ws.freeze_panes = "A2"
        
        return ws
    
    def _create_rekap_poli_sheet(self, wb, df_grid, slot_str):
        """Buat sheet Rekap Poli"""
        if "Rekap Poli" in wb.sheetnames:
            del wb["Rekap Poli"]
        
        ws = wb.create_sheet("Rekap Poli")
        ws.append(["POLI", "HARI", "REGULER (JAM)", "POLEKS (JAM)", "TOTAL JAM"])
        
        if df_grid is not None and not df_grid.empty:
            for (poli, hari), group in df_grid.groupby(["POLI", "HARI"]):
                total_r = total_e = 0
                
                for slot in slot_str:
                    if slot in group.columns:
                        total_r += (group[slot] == "R").sum().sum()
                        total_e += (group[slot] == "E").sum().sum()
                
                # Convert to hours
                hours_r = round(total_r * self.interval / 60, 2)
                hours_e = round(total_e * self.interval / 60, 2)
                total_hours = hours_r + hours_e
                
                ws.append([poli, hari, hours_r, hours_e, total_hours])
        
        # Add totals row
        if ws.max_row > 1:
            ws.append(["TOTAL", "", 
                      f"=SUM(C2:C{ws.max_row})", 
                      f"=SUM(D2:D{ws.max_row})", 
                      f"=SUM(E2:E{ws.max_row})"])
        
        # Style
        self._style_rekap_sheet(ws)
        ws.freeze_panes = "A2"
        
        return ws
    
    def _create_rekap_dokter_sheet(self, wb, df_grid, slot_str):
        """Buat sheet Rekap Dokter"""
        if "Rekap Dokter" in wb.sheetnames:
            del wb["Rekap Dokter"]
        
        ws = wb.create_sheet("Rekap Dokter")
        ws.append(["DOKTER", "HARI", "SHIFT", "TOTAL JAM"])
        
        if df_grid is not None and not df_grid.empty:
            for (dokter, hari), group in df_grid.groupby(["DOKTER", "HARI"]):
                active_slots = [s for s in slot_str if s in group.columns and group.iloc[0].get(s) in ["R", "E"]]
                
                if active_slots:
                    time_ranges = self._combine_slots_to_ranges(active_slots)
                    
                    for time_range in time_ranges:
                        # Calculate duration
                        duration = self._calculate_duration(time_range, slot_str)
                        ws.append([dokter, hari, time_range, round(duration, 2)])
        
        # Style
        self._style_rekap_sheet(ws)
        ws.freeze_panes = "A2"
        
        return ws
    
    def _create_peak_hour_sheet(self, wb, df_grid, slot_str):
        """Buat sheet Peak Hour Analysis"""
        if "Peak Hour Analysis" in wb.sheetnames:
            del wb["Peak Hour Analysis"]
        
        ws = wb.create_sheet("Peak Hour Analysis")
        ws.append(["HARI", "SLOT", "JUMLAH DOKTER", "LEVEL"])
        
        if df_grid is not None and not df_grid.empty:
            for hari, group in df_grid.groupby("HARI"):
                slot_counts = []
                
                for slot in slot_str:
                    if slot in group.columns:
                        count = ((group[slot] == "R") | (group[slot] == "E")).sum()
                        slot_counts.append((slot, count))
                
                if slot_counts:
                    max_count = max(count for _, count in slot_counts)
                    
                    for slot, count in slot_counts:
                        if count == max_count:
                            level = "HIGH"
                        elif count >= max_count * 0.7:
                            level = "MEDIUM"
                        else:
                            level = "LOW"
                        
                        ws.append([hari, slot, count, level])
        
        # Style
        self._style_rekap_sheet(ws)
        ws.freeze_panes = "A2"
        
        return ws
    
    def _create_conflict_doctor_sheet(self, wb, df_grid, slot_str):
        """Buat sheet Conflict Dokter"""
        if "Conflict Dokter" in wb.sheetnames:
            del wb["Conflict Dokter"]
        
        ws = wb.create_sheet("Conflict Dokter")
        ws.append(["DOKTER", "HARI", "SLOT", "KETERANGAN", "TINGKAT"])
        
        if df_grid is not None and not df_grid.empty:
            conflicts = self._find_doctor_conflicts(df_grid, slot_str)
            
            for conflict in conflicts:
                ws.append([
                    conflict["dokter"],
                    conflict["hari"],
                    conflict["slot"],
                    conflict["keterangan"],
                    conflict["tingkat"]
                ])
        
        # Style
        self._style_conflict_sheet(ws)
        ws.freeze_panes = "A2"
        
        return ws
    
    def _create_conflict_map_sheet(self, wb, df_grid, slot_str):
        """Buat sheet Peta Konflik Dokter"""
        if "Peta Konflik Dokter" in wb.sheetnames:
            del wb["Peta Konflik Dokter"]
        
        ws = wb.create_sheet("Peta Konflik Dokter")
        
        # Get unique doctors
        doctors = sorted(df_grid["DOKTER"].unique()) if df_grid is not None else []
        
        # Header row
        header = ["SLOT"] + doctors
        ws.append(header)
        
        # Data rows
        for slot in slot_str:
            row = [slot]
            for doctor in doctors:
                row.append("")  # Placeholder
            ws.append(row)
        
        # Fill conflict data
        if df_grid is not None and not df_grid.empty:
            for (dokter, hari), group in df_grid.groupby(["DOKTER", "HARI"]):
                col_idx = doctors.index(dokter) + 2  # +1 for header, +1 for 1-based indexing
                
                for slot in slot_str:
                    if slot in group.columns:
                        row_idx = slot_str.index(slot) + 2  # +1 for header, +1 for 1-based indexing
                        cell = ws.cell(row=row_idx, column=col_idx)
                        
                        values = group[slot].unique()
                        
                        if len(values) > 1 and any(v in ["R", "E"] for v in values):
                            cell.value = "⚠️"
                            cell.fill = self.fill_conflict
                        
                        if "R" in values and "E" in values:
                            cell.value = "🚨"
                            cell.fill = self.fill_conflict_hard
        
        # Style
        self._style_conflict_map_sheet(ws)
        
        return ws
    
    def _create_grafik_poli_sheet(self, wb, df_grid):
        """Buat sheet Grafik Poli"""
        if "Grafik Poli" in wb.sheetnames:
            del wb["Grafik Poli"]
        
        ws = wb.create_sheet("Grafik Poli")
        
        # Try to get data from Rekap Poli sheet
        try:
            if "Rekap Poli" in wb.sheetnames:
                rp_ws = wb["Rekap Poli"]
                
                # Collect poli totals
                poli_totals = {}
                for row in rp_ws.iter_rows(min_row=2, max_col=5, values_only=True):
                    if row and row[0] and row[0] != "TOTAL":
                        poli = row[0]
                        total = row[4] if row[4] is not None else 0
                        poli_totals[poli] = poli_totals.get(poli, 0) + total
                
                # Write data
                ws.append(["POLI", "TOTAL JAM"])
                for poli, total in sorted(poli_totals.items(), key=lambda x: x[1], reverse=True):
                    ws.append([poli, total])
                
                # Create chart
                if len(poli_totals) > 0:
                    chart = BarChart()
                    chart.title = "Beban Poli (Total Jam)"
                    chart.style = 10
                    chart.y_axis.title = "Total Jam"
                    chart.x_axis.title = "Poli"
                    
                    data = Reference(ws, min_col=2, min_row=1, max_row=ws.max_row)
                    categories = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)
                    
                    chart.add_data(data, titles_from_data=True)
                    chart.set_categories(categories)
                    
                    ws.add_chart(chart, "E5")
        
        except Exception as e:
            print(f"⚠️ Could not create Grafik Poli: {e}")
            ws.append(["POLI", "TOTAL JAM"])
            ws.append(["Data tidak tersedia", 0])
        
        # Style
        self._style_chart_sheet(ws)
        
        return ws
    
    def _create_summary_sheet(self, wb, df_grid, slot_str):
        """Buat sheet Summary dengan statistik"""
        if "Summary" in wb.sheetnames:
            del wb["Summary"]
        
        ws = wb.create_sheet("Summary")
        
        # Title
        ws.merge_cells("A1:F1")
        title_cell = ws["A1"]
        title_cell.value = "SUMMARY JADWAL DOKTER"
        title_cell.font = Font(bold=True, size=16, color="366092")
        title_cell.alignment = Alignment(horizontal="center")
        
        # Statistics
        stats = self._calculate_statistics(df_grid, slot_str)
        
        ws.append([])  # Empty row
        ws.append(["STATISTIK", ""])
        ws.append(["Total Baris Data", stats["total_rows"]])
        ws.append(["Total Dokter Unik", stats["total_doctors"]])
        ws.append(["Total Poli Unik", stats["total_poli"]])
        ws.append(["Total Slot Waktu", stats["total_slots"]])
        ws.append(["Slot Reguler (R)", stats["total_r"]])
        ws.append(["Slot Poleks (E)", stats["total_e"]])
        ws.append(["Slot Kosong", stats["total_empty"]])
        ws.append(["Persentase Terisi", f"{stats['fill_percentage']:.1f}%"])
        
        # Configuration
        ws.append([])  # Empty row
        ws.append(["KONFIGURASI", ""])
        ws.append(["Jam Mulai", f"{self.config.start_hour:02d}:{self.config.start_minute:02d}"])
        ws.append(["Interval Slot", f"{self.config.interval_minutes} menit"])
        ws.append(["Maks Poleks/Slot", self.config.max_poleks_per_slot])
        ws.append(["Auto Fix Errors", "Ya" if self.config.auto_fix_errors else "Tidak"])
        ws.append(["Hari Sabtu", "Aktif" if self.config.enable_sabtu else "Nonaktif"])
        
        # Sheet list
        ws.append([])  # Empty row
        ws.append(["DAFTAR SHEET", ""])
        for i, sheet_name in enumerate(wb.sheetnames, 1):
            ws.append([f"{i}.", sheet_name])
        
        # Timestamp
        ws.append([])  # Empty row
        ws.append(["Dibuat pada", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
        
        # Style
        self._style_summary_sheet(ws)
        
        return ws
    
    # ======================================================
    # STYLING METHODS
    # ======================================================
    
    def _style_jadwal_sheet(self, ws, df_grid, slot_str):
        """Style sheet Jadwal"""
        # Style header
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col)
            cell.fill = self.fill_header
            cell.font = self.font_header
            cell.alignment = self.align_center
            cell.border = self.thin_border
        
        # Apply alternating row colors
        for row in range(2, ws.max_row + 1):
            # Style metadata columns
            for col in range(1, 6):  # Columns A-E (POLI to JAM)
                cell = ws.cell(row=row, column=col)
                cell.border = self.thin_border
                if col <= 4:  # POLI to HARI
                    cell.alignment = self.align_left
                else:  # JAM
                    cell.alignment = self.align_center
            
            # Alternate row background
            if row % 2 == 0:
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col).fill = self.fill_gray
        
        # Color code time slots
        if df_grid is not None and not df_grid.empty:
            poleks_counter = {}
            
            for row_idx, row_data in enumerate(df_grid.itertuples(index=False), start=2):
                hari = getattr(row_data, 'HARI', '')
                
                for col_idx, slot in enumerate(slot_str, start=6):  # Start from column F
                    if hasattr(row_data, slot):
                        value = getattr(row_data, slot)
                        cell = ws.cell(row=row_idx, column=col_idx)
                        
                        if value == "R":
                            cell.fill = self.fill_r
                        elif value == "E":
                            # Count poleks per hari per slot
                            key = (hari, slot)
                            poleks_counter[key] = poleks_counter.get(key, 0) + 1
                            
                            if poleks_counter[key] > self.max_e:
                                cell.fill = self.fill_over
                            else:
                                cell.fill = self.fill_e
                        
                        cell.alignment = self.align_center
                        cell.border = self.thin_border
    
    def _style_rekap_sheet(self, ws):
        """Style sheet rekap"""
        # Header
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col)
            cell.fill = self.fill_header
            cell.font = self.font_header
            cell.alignment = self.align_center
            cell.border = self.thin_border
        
        # Data rows
        for row in range(2, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=row, column=col)
                cell.border = self.thin_border
                
                # Align numeric columns to right
                if col >= 3:  # Assume column 3+ are numeric
                    try:
                        float(cell.value)
                        cell.alignment = self.align_right
                        cell.number_format = '#,##0.00'
                    except:
                        cell.alignment = self.align_left
                else:
                    cell.alignment = self.align_left
            
            # Alternate row colors
            if row % 2 == 0:
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col).fill = self.fill_gray
        
        # Style total row if exists
        if ws.max_row > 2:
            last_row = ws.max_row
            if "TOTAL" in str(ws.cell(row=last_row, column=1).value):
                for col in range(1, ws.max_column + 1):
                    cell = ws.cell(row=last_row, column=col)
                    cell.fill = self.fill_total
                    cell.font = self.font_bold
    
    def _style_conflict_sheet(self, ws):
        """Style conflict sheet"""
        self._style_rekap_sheet(ws)  # Use same styling
        
        # Highlight conflict rows
        for row in range(2, ws.max_row + 1):
            tingkat = ws.cell(row=row, column=5).value  # Column E = TINGKAT
            if tingkat == "TINGGI":
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col).fill = PatternFill("solid", fgColor="FFC7CE")
    
    def _style_conflict_map_sheet(self, ws):
        """Style conflict map sheet"""
        # Header
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col)
            cell.fill = self.fill_header
            cell.font = self.font_header
            cell.alignment = self.align_center
            cell.border = self.thick_border
        
        # Slot column
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=1)
            cell.font = self.font_bold
            cell.alignment = self.align_center
            cell.border = self.thin_border
        
        # Doctor columns
        for row in range(2, ws.max_row + 1):
            for col in range(2, ws.max_column + 1):
                cell = ws.cell(row=row, column=col)
                cell.alignment = self.align_center
                cell.border = self.thin_border
    
    def _style_chart_sheet(self, ws):
        """Style chart sheet"""
        # Header
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col)
            cell.fill = self.fill_header
            cell.font = self.font_header
            cell.alignment = self.align_center
            cell.border = self.thin_border
        
        # Data
        for row in range(2, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=row, column=col)
                cell.border = self.thin_border
                cell.alignment = self.align_left if col == 1 else self.align_right
        
        # Alternate rows
        for row in range(2, ws.max_row + 1):
            if row % 2 == 0:
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col).fill = self.fill_gray
    
    def _style_summary_sheet(self, ws):
        """Style summary sheet"""
        # Make all cells have borders
        for row in range(1, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=row, column=col)
                cell.border = self.thin_border
        
        # Style headers
        header_rows = [2, 10, 17]  # Row numbers with section headers
        for row in header_rows:
            if row <= ws.max_row:
                for col in range(1, 3):
                    cell = ws.cell(row=row, column=col)
                    cell.fill = PatternFill("solid", fgColor="4F81BD")
                    cell.font = Font(bold=True, color="FFFFFF", size=11)
        
        # Style data rows
        for row in range(1, ws.max_row + 1):
            # Left align first column, right align second column
            cell1 = ws.cell(row=row, column=1)
            cell2 = ws.cell(row=row, column=2)
            
            cell1.alignment = self.align_left
            cell2.alignment = self.align_left if row in [1, 19] else self.align_right
            
            # Bold important rows
            if row in [3, 4, 5, 8, 9]:
                cell1.font = self.font_bold
                cell2.font = self.font_bold
    
    def _apply_styling_to_all_sheets(self, wb):
        """Apply basic styling to all sheets"""
        for ws in wb.worksheets:
            # Set default font
            for row in ws.iter_rows():
                for cell in row:
                    if cell.font is None or cell.font.name == 'Calibri':
                        cell.font = self.font_normal
    
    def _auto_adjust_column_widths(self, wb):
        """Auto adjust column widths for all sheets"""
        for ws in wb.worksheets:
            for column in ws.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)
                
                for cell in column:
                    try:
                        if cell.value:
                            cell_length = len(str(cell.value))
                            if cell_length > max_length:
                                max_length = cell_length
                    except:
                        pass
                
                adjusted_width = min(max_length + 2, 50)  # Max width 50
                ws.column_dimensions[column_letter].width = adjusted_width
    
    # ======================================================
    # HELPER METHODS
    # ======================================================
    
    def _load_workbook(self, source_file):
        """Load workbook dari berbagai sumber"""
        try:
            if hasattr(source_file, 'read'):
                # BytesIO atau file-like object
                source_file.seek(0)
                return load_workbook(source_file)
            elif isinstance(source_file, str):
                # File path
                return load_workbook(source_file)
            else:
                # Buat workbook baru
                wb = Workbook()
                if "Sheet" in wb.sheetnames:
                    del wb["Sheet"]
                return wb
        except Exception as e:
            print(f"⚠️ Could not load workbook: {e}")
            return None
    
    def _combine_slots_to_ranges(self, slots):
        """Gabungkan slot menjadi range waktu"""
        if not slots:
            return []
        
        try:
            # Parse slot times
            slot_times = []
            for slot in slots:
                try:
                    dt = datetime.strptime(slot, "%H:%M")
                    slot_times.append((slot, dt))
                except:
                    pass
            
            if not slot_times:
                return slots
            
            # Sort by time
            slot_times.sort(key=lambda x: x[1])
            
            ranges = []
            current_start = slot_times[0][0]
            current_end = slot_times[0][0]
            
            for i in range(1, len(slot_times)):
                current_time = slot_times[i][1]
                prev_time = datetime.strptime(current_end, "%H:%M")
                
                # Check if consecutive (within interval)
                time_diff = (current_time - prev_time).seconds / 60
                
                if time_diff == self.interval:
                    current_end = slot_times[i][0]
                else:
                    if current_start == current_end:
                        ranges.append(current_start)
                    else:
                        ranges.append(f"{current_start}-{current_end}")
                    
                    current_start = slot_times[i][0]
                    current_end = slot_times[i][0]
            
            # Add last range
            if current_start == current_end:
                ranges.append(current_start)
            else:
                ranges.append(f"{current_start}-{current_end}")
            
            return ranges
            
        except Exception as e:
            print(f"⚠️ Error combining slots: {e}")
            return slots
    
    def _calculate_duration(self, time_range, slot_str):
        """Hitung durasi dalam jam dari range waktu"""
        try:
            if '-' in time_range:
                start_str, end_str = time_range.split('-')
                
                if start_str in slot_str and end_str in slot_str:
                    start_idx = slot_str.index(start_str)
                    end_idx = slot_str.index(end_str)
                    num_slots = end_idx - start_idx + 1
                    return num_slots * self.interval / 60
            
            # Fallback: single slot
            return self.interval / 60
            
        except:
            return 0
    
    def _find_doctor_conflicts(self, df_grid, slot_str):
        """Temukan konflik dokter"""
        conflicts = []
        
        if df_grid is None or df_grid.empty:
            return conflicts
        
        for (dokter, hari), group in df_grid.groupby(["DOKTER", "HARI"]):
            if len(group) > 1:  # Dokter muncul di >1 poli di hari yang sama
                for slot in slot_str:
                    if slot in group.columns:
                        active_polis = []
                        
                        for _, row in group.iterrows():
                            if row[slot] in ["R", "E"]:
                                active_polis.append(row["POLI"])
                        
                        if len(active_polis) > 1:
                            # Check for R vs E conflict
                            has_r = any(group[slot] == "R")
                            has_e = any(group[slot] == "E")
                            
                            if has_r and has_e:
                                tingkat = "TINGGI"
                                keterangan = f"Bentrok Reguler & Poleks di {', '.join(active_polis)}"
                            else:
                                tingkat = "SEDANG"
                                keterangan = f"Poli berbeda di waktu sama: {', '.join(active_polis)}"
                            
                            conflicts.append({
                                "dokter": dokter,
                                "hari": hari,
                                "slot": slot,
                                "keterangan": keterangan,
                                "tingkat": tingkat
                            })
        
        return conflicts
    
    def _calculate_statistics(self, df_grid, slot_str):
        """Hitung statistik untuk summary"""
        stats = {
            "total_rows": 0,
            "total_doctors": 0,
            "total_poli": 0,
            "total_slots": 0,
            "total_r": 0,
            "total_e": 0,
            "total_empty": 0,
            "fill_percentage": 0
        }
        
        if df_grid is not None and not df_grid.empty:
            stats["total_rows"] = len(df_grid)
            stats["total_doctors"] = df_grid["DOKTER"].nunique()
            stats["total_poli"] = df_grid["POLI"].nunique()
            
            total_cells = len(df_grid) * len(slot_str)
            stats["total_slots"] = total_cells
            
            for slot in slot_str:
                if slot in df_grid.columns:
                    stats["total_r"] += (df_grid[slot] == "R").sum()
                    stats["total_e"] += (df_grid[slot] == "E").sum()
            
            stats["total_empty"] = total_cells - stats["total_r"] - stats["total_e"]
            
            if total_cells > 0:
                stats["fill_percentage"] = ((stats["total_r"] + stats["total_e"]) / total_cells) * 100
        
        return stats
    
    # ======================================================
    # TEMPLATE GENERATOR
    # ======================================================
    
    def generate_template(self, slot_str=None):
        """
        Generate template Excel file untuk input
        
        Args:
            slot_str: List slot waktu (opsional)
            
        Returns:
            BytesIO buffer berisi template Excel
        """
        print("📄 Generating template Excel...")
        
        try:
            wb = Workbook()
            
            # Remove default sheet
            if "Sheet" in wb.sheetnames:
                del wb["Sheet"]
            
            # Create Reguler sheet
            ws_reg = wb.create_sheet("Reguler")
            self._create_template_sheet(ws_reg, "Reguler", slot_str)
            
            # Create Poleks sheet
            ws_pol = wb.create_sheet("Poleks")
            self._create_template_sheet(ws_pol, "Poleks", slot_str)
            
            # Create Poli Asal sheet
            ws_poli = wb.create_sheet("Poli Asal")
            self._create_poli_asal_sheet(ws_poli)
            
            # Create Instructions sheet
            ws_inst = wb.create_sheet("Instruksi")
            self._create_instructions_sheet(ws_inst)
            
            # Apply styling
            self._apply_styling_to_all_sheets(wb)
            self._auto_adjust_column_widths(wb)
            
            # Save to buffer
            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            
            print(f"✅ Template created: {buf.getbuffer().nbytes} bytes")
            return buf
            
        except Exception as e:
            print(f"❌ Error generating template: {e}")
            print(traceback.format_exc())
            raise
    
    def _create_template_sheet(self, ws, jenis, slot_str):
        """Buat sheet template untuk Reguler atau Poleks"""
        # Header
        headers = ["Nama Dokter", "Poli Asal", "Jenis Poli", "Senin", "Selasa", "Rabu", "Kamis", "Jum'at"]
        
        if self.config.enable_sabtu:
            headers.append("Sabtu")
        
        ws.append(headers)
        
        # Example data
        examples = [
            ["dr. Contoh Satu", "Poli Anak", jenis, "08.00-10.00", "09.00-11.00", "", "10.00-12.00", "08.00-10.00"],
            ["dr. Contoh Dua", "Poli Dalam", jenis, "10.00-12.00", "", "08.00-10.00", "09.00-11.00", ""],
        ]
        
        if self.config.enable_sabtu:
            for ex in examples:
                ex.append("")  # Add empty Sabtu column
        
        for example in examples:
            ws.append(example)
        
        # Notes row
        ws.append([])
        ws.append(["CATATAN:"])
        ws.append(["1. Format waktu: '07.30-10.00' atau '07:30-10:00'"])
        ws.append(["2. Kosongkan jika tidak ada jadwal"])
        ws.append(["3. Jenis Poli otomatis diisi berdasarkan nama sheet"])
        
        # Style
        for col in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col)
            cell.fill = self.fill_header
            cell.font = self.font_header
            cell.alignment = self.align_center
        
        # Style example rows
        for row in range(2, 4):
            for col in range(1, len(headers) + 1):
                cell = ws.cell(row=row, column=col)
                cell.border = self.thin_border
                cell.alignment = self.align_left
        
        # Style notes
        notes_row = len(examples) + 2
        ws.merge_cells(f"A{notes_row}:{get_column_letter(len(headers))}{notes_row}")
        note_cell = ws.cell(row=notes_row, column=1)
        note_cell.font = Font(bold=True, color="FF0000")
        
        ws.freeze_panes = "D2"  # Freeze header and first 3 columns
    
    def _create_poli_asal_sheet(self, ws):
        """Buat sheet referensi Poli Asal"""
        ws.append(["No", "Nama Poli", "Kode Sheet"])
        
        polis = [
            [1, "Poli Anak", "ANAK"],
            [2, "Poli Bedah", "BEDAH"],
            [3, "Poli Dalam", "DALAM"],
            [4, "Poli Obgyn", "OBGYN"],
            [5, "Poli Jantung", "JANTUNG"],
            [6, "Poli Ortho", "ORTHO"],
            [7, "Poli Paru", "PARU"],
            [8, "Poli Saraf", "SARAF"],
            [9, "Poli THT", "THT"],
            [10, "Poli Urologi", "URO"],
            [11, "Poli Jiwa", "JIWA"],
            [12, "Poli Kukel", "KUKEL"],
            [13, "Poli Bedah Saraf", "BSARAF"],
            [14, "Poli Gigi", "GIGI"],
            [15, "Poli Mata", "MATA"],
            [16, "Poli Rehab", "REHAB"],
        ]
        
        for poli in polis:
            ws.append(poli)
        
        # Style
        for col in range(1, 4):
            cell = ws.cell(row=1, column=col)
            cell.fill = self.fill_header
            cell.font = self.font_header
            cell.alignment = self.align_center
        
        for row in range(2, len(polis) + 2):
            for col in range(1, 4):
                cell = ws.cell(row=row, column=col)
                cell.border = self.thin_border
                cell.alignment = self.align_left
        
        ws.freeze_panes = "A2"
    
    def _create_instructions_sheet(self, ws):
        """Buat sheet instruksi"""
        title = ws.cell(row=1, column=1)
        title.value = "PETUNJUK PENGGUNAAN"
        title.font = Font(bold=True, size=14, color="366092")
        
        instructions = [
            "",
            "1. FILE INPUT:",
            "   - File Excel harus memiliki 2 sheet: 'Reguler' dan 'Poleks'",
            "   - Format kolom harus sama seperti di sheet template",
            "",
            "2. FORMAT WAKTU:",
            "   - Gunakan format: '07.30-10.00' atau '07:30-10:00'",
            "   - Bisa menggunakan titik atau titik dua sebagai separator",
            "   - Kosongkan sel jika tidak ada jadwal",
            "",
            "3. KOLOM WAJIB:",
            "   - Nama Dokter: Nama lengkap dokter",
            "   - Poli Asal: Nama poli (lihat sheet 'Poli Asal' untuk referensi)",
            "   - Jenis Poli: Akan otomatis terisi berdasarkan nama sheet",
            "   - Kolom hari: Senin sampai Jum'at (Sabtu optional)",
            "",
            "4. PROSES:",
            "   - Upload file di tab 'Upload & Proses'",
            "   - Klik 'Proses Jadwal' untuk konversi ke format grid",
            "   - Download hasil di file Excel lengkap",
            "",
            "5. OUTPUT:",
            "   - File hasil akan berisi 10+ sheet dengan analisis lengkap",
            "   - Termasuk deteksi konflik dan statistik",
            "",
            "6. PENGATURAN:",
            "   - Ubah pengaturan di sidebar jika perlu",
            "   - Atur jam mulai, interval, dan batasan poleks",
            "",
            "© 2024 Sistem Jadwal Dokter"
        ]
        
        for i, instruction in enumerate(instructions, start=3):
            cell = ws.cell(row=i, column=1)
            cell.value = instruction
            if instruction.startswith(("1.", "2.", "3.", "4.", "5.", "6.")):
                cell.font = Font(bold=True)
        
        # Adjust column width
        ws.column_dimensions['A'].width = 100
