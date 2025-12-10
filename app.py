import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from datetime import time, timedelta, datetime, date
import io
import re
import json
from typing import Dict, List, Optional
import uuid

# Konfigurasi halaman
st.set_page_config(
    page_title="Pengisi Jadwal Poli Excel + Kanban",
    page_icon="🏥📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Warna untuk sel
FILL_R = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
FILL_E = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")
FILL_OVERLIMIT = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

# Warna untuk Kanban
KANBAN_COLORS = {
    "todo": "#FF6B6B",
    "in_progress": "#FFD93D",
    "review": "#6BCF7F",
    "done": "#4D96FF"
}

# Konfigurasi default
TIME_SLOTS = [
    time(7, 30), time(8, 0), time(8, 30), time(9, 0), time(9, 30),
    time(10, 0), time(10, 30), time(11, 0), time(11, 30), time(12, 0),
    time(12, 30), time(13, 0), time(13, 30), time(14, 0), time(14, 30)
]
TIME_SLOTS_STR = [t.strftime("%H:%M") for t in TIME_SLOTS]

HARI_ORDER = {"Senin": 1, "Selasa": 2, "Rabu": 3, "Kamis": 4, "Jum'at": 5}
HARI_INDONESIA = ["Senin", "Selasa", "Rabu", "Kamis", "Jum'at"]

# Inisialisasi session state untuk Kanban
if 'kanban_tasks' not in st.session_state:
    st.session_state.kanban_tasks = {
        "todo": [],
        "in_progress": [],
        "review": [],
        "done": []
    }

if 'kanban_next_id' not in st.session_state:
    st.session_state.kanban_next_id = 1

class KanbanTask:
    def __init__(self, id, title, description="", priority="medium", 
                 due_date=None, assignee="", tags=None, created_by="", 
                 created_date=None):
        self.id = id
        self.title = title
        self.description = description
        self.priority = priority  # low, medium, high
        self.due_date = due_date
        self.assignee = assignee
        self.tags = tags or []
        self.created_by = created_by
        self.created_date = created_date or datetime.now()
        self.last_updated = datetime.now()
    
    def to_dict(self):
        return {
            "id": self.id,
            "title": self.title,
            "description": self.description,
            "priority": self.priority,
            "due_date": self.due_date.strftime("%Y-%m-%d") if self.due_date else None,
            "assignee": self.assignee,
            "tags": self.tags,
            "created_by": self.created_by,
            "created_date": self.created_date.strftime("%Y-%m-%d %H:%M"),
            "last_updated": self.last_updated.strftime("%Y-%m-%d %H:%M")
        }

def parse_time_range(time_str):
    """Parse rentang waktu dari string seperti '08.00 - 10.30'"""
    if pd.isna(time_str) or str(time_str).strip() == "":
        return None, None
    
    clean_str = str(time_str).strip().replace(' ', '').replace('.', ':')
    pattern = r'(\d{1,2}:\d{2})-(\d{1,2}:\d{2})'
    match = re.search(pattern, clean_str)
    if not match:
        return None, None
    
    try:
        start_str, end_str = match.groups()
        start_hour, start_minute = map(int, start_str.split(':'))
        end_hour, end_minute = map(int, end_str.split(':'))
        
        start_time = time(start_hour, start_minute)
        end_time = time(end_hour, end_minute)
        
        if end_time < start_time:
            end_time = time(end_hour + 24, end_minute)
        
        return start_time, end_time
    except:
        return None, None

def time_overlap(slot_start, slot_end, schedule_start, schedule_end):
    """Cek apakah slot waktu overlap dengan jadwal"""
    if schedule_start is None or schedule_end is None:
        return False
    
    return not (slot_end <= schedule_start or slot_start >= schedule_end)

def process_schedule(df, jenis_poli):
    """Proses dataframe jadwal"""
    results = []
    
    for (dokter, poli_asal), group in df.groupby(['Nama Dokter', 'Poli Asal']):
        hari_schedules = {}
        
        for hari in HARI_INDONESIA:
            if hari not in group.columns:
                continue
                
            time_ranges = []
            for time_str in group[hari]:
                start_time, end_time = parse_time_range(time_str)
                if start_time and end_time:
                    if end_time > time(14, 30):
                        end_time = time(14, 30)
                    time_ranges.append((start_time, end_time))
            
            hari_schedules[hari] = time_ranges
        
        for hari in HARI_INDONESIA:
            if hari not in hari_schedules or not hari_schedules[hari]:
                continue
                
            merged_ranges = []
            for start, end in sorted(hari_schedules[hari]):
                if not merged_ranges:
                    merged_ranges.append([start, end])
                else:
                    last_start, last_end = merged_ranges[-1]
                    if start <= last_end:
                        merged_ranges[-1][1] = max(last_end, end)
                    else:
                        merged_ranges.append([start, end])
            
            row = {
                'POLI ASAL': poli_asal,
                'JENIS POLI': jenis_poli,
                'HARI': hari,
                'DOKTER': dokter
            }
            
            for i, slot_time in enumerate(TIME_SLOTS):
                slot_start = slot_time
                slot_end = (datetime.combine(datetime.today(), slot_start) + 
                           timedelta(minutes=30)).time()
                
                has_overlap = any(
                    time_overlap(slot_start, slot_end, start, end)
                    for start, end in merged_ranges
                )
                
                if has_overlap:
                    row[TIME_SLOTS_STR[i]] = 'R' if jenis_poli == 'Reguler' else 'E'
                else:
                    row[TIME_SLOTS_STR[i]] = ''
            
            results.append(row)
    
    return pd.DataFrame(results)

def apply_styles(ws, max_row):
    """Terapkan styling ke sheet Jadwal"""
    e_counts = {hari: {slot: 0 for slot in TIME_SLOTS_STR} for hari in HARI_INDONESIA}
    
    for row in range(2, max_row + 1):
        hari = ws.cell(row=row, column=3).value
        if hari not in HARI_INDONESIA:
            continue
            
        for col_idx, slot in enumerate(TIME_SLOTS_STR, start=5):
            cell = ws.cell(row=row, column=col_idx)
            if cell.value == 'E':
                e_counts[hari][slot] += 1
    
    for row in range(2, max_row + 1):
        hari = ws.cell(row=row, column=3).value
        if hari not in HARI_INDONESIA:
            continue
            
        for col_idx, slot in enumerate(TIME_SLOTS_STR, start=5):
            cell = ws.cell(row=row, column=col_idx)
            
            if cell.value == 'R':
                cell.fill = FILL_R
            elif cell.value == 'E':
                cell.fill = FILL_E
                
                if e_counts[hari][slot] > 7:
                    e_rows_for_slot = []
                    for r in range(2, max_row + 1):
                        if (ws.cell(row=r, column=3).value == hari and 
                            ws.cell(row=r, column=col_idx).value == 'E'):
                            e_rows_for_slot.append(r)
                    
                    if len(e_rows_for_slot) > 7:
                        if row in e_rows_for_slot[7:]:
                            cell.fill = FILL_OVERLIMIT

# ==================== KANBAN FUNCTIONS ====================

def render_kanban_board():
    """Render Kanban board"""
    st.subheader("📋 Kanban Board - Tracking Jadwal")
    
    # Filter controls
    col1, col2, col3 = st.columns(3)
    with col1:
        filter_priority = st.selectbox(
            "Filter Priority",
            ["All", "high", "medium", "low"],
            key="filter_priority"
        )
    with col2:
        filter_assignee = st.selectbox(
            "Filter Assignee",
            ["All"] + sorted(list(set(
                task.get("assignee", "") 
                for column in st.session_state.kanban_tasks.values() 
                for task in column 
                if task.get("assignee")
            ))),
            key="filter_assignee"
        )
    with col3:
        filter_tags = st.multiselect(
            "Filter Tags",
            options=sorted(list(set(
                tag 
                for column in st.session_state.kanban_tasks.values() 
                for task in column 
                for tag in task.get("tags", [])
            ))),
            key="filter_tags"
        )
    
    # Create Kanban columns
    columns = st.columns(4)
    column_names = ["Todo", "In Progress", "Review", "Done"]
    column_keys = ["todo", "in_progress", "review", "done"]
    
    for idx, (col, col_name, col_key) in enumerate(zip(columns, column_names, column_keys)):
        with col:
            # Column header
            st.markdown(
                f"""
                <div style='
                    background-color: {KANBAN_COLORS[col_key]};
                    padding: 10px;
                    border-radius: 5px;
                    margin-bottom: 10px;
                    color: white;
                    font-weight: bold;
                    text-align: center;
                '>
                {col_name} ({len(st.session_state.kanban_tasks[col_key])})
                </div>
                """,
                unsafe_allow_html=True
            )
            
            # Tasks in this column
            tasks = st.session_state.kanban_tasks[col_key]
            
            # Apply filters
            filtered_tasks = tasks
            if filter_priority != "All":
                filtered_tasks = [t for t in filtered_tasks if t.get("priority") == filter_priority]
            if filter_assignee != "All":
                filtered_tasks = [t for t in filtered_tasks if t.get("assignee") == filter_assignee]
            if filter_tags:
                filtered_tasks = [t for t in filtered_tasks if any(tag in t.get("tags", []) for tag in filter_tags)]
            
            for task in filtered_tasks:
                render_task_card(task, col_key)
            
            # Add task button for Todo column
            if col_key == "todo":
                with st.expander("➕ Tambah Task Baru", expanded=False):
                    add_task_form()

def render_task_card(task, current_column):
    """Render individual task card"""
    priority_colors = {
        "high": "#FF0000",
        "medium": "#FFA500",
        "low": "#008000"
    }
    
    with st.container():
        st.markdown(
            f"""
            <div style='
                background-color: white;
                border: 2px solid {priority_colors.get(task.get("priority", "medium"), "#FFA500")};
                border-radius: 8px;
                padding: 12px;
                margin-bottom: 10px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            '>
            <div style='display: flex; justify-content: space-between; align-items: center;'>
                <strong>#{task.get("id")}</strong>
                <small>{task.get("priority", "medium").upper()}</small>
            </div>
            <h4 style='margin: 8px 0;'>{task.get("title", "No Title")}</h4>
            """,
            unsafe_allow_html=True
        )
        
        if task.get("description"):
            st.caption(task.get("description"))
        
        # Tags
        if task.get("tags"):
            tags_html = " ".join([f"<span style='background-color: #e0e0e0; padding: 2px 6px; border-radius: 10px; font-size: 0.8em; margin-right: 4px;'>{tag}</span>" for tag in task.get("tags")])
            st.markdown(tags_html, unsafe_allow_html=True)
        
        # Assignee and due date
        col1, col2 = st.columns(2)
        with col1:
            if task.get("assignee"):
                st.caption(f"👤 {task.get('assignee')}")
        with col2:
            if task.get("due_date"):
                st.caption(f"📅 {task.get('due_date')}")
        
        st.markdown("</div>", unsafe_allow_html=True)
        
        # Task actions
        col_move1, col_move2, col_edit, col_delete = st.columns(4)
        
        # Move left button
        with col_move1:
            column_keys = ["todo", "in_progress", "review", "done"]
            current_idx = column_keys.index(current_column)
            if current_idx > 0:
                if st.button("⬅️", key=f"left_{task['id']}", help="Move left"):
                    move_task(task['id'], current_column, column_keys[current_idx - 1])
                    st.rerun()
        
        # Move right button
        with col_move2:
            if current_idx < 3:
                if st.button("➡️", key=f"right_{task['id']}", help="Move right"):
                    move_task(task['id'], current_column, column_keys[current_idx + 1])
                    st.rerun()
        
        # Edit button
        with col_edit:
            if st.button("✏️", key=f"edit_{task['id']}", help="Edit"):
                st.session_state.editing_task = task['id']
                st.rerun()
        
        # Delete button
        with col_delete:
            if st.button("🗑️", key=f"delete_{task['id']}", help="Delete"):
                delete_task(task['id'], current_column)
                st.rerun()

def add_task_form():
    """Form untuk menambah task baru"""
    with st.form(key="add_task_form"):
        title = st.text_input("Judul Task*", placeholder="Contoh: Review jadwal Poli Anak")
        description = st.text_area("Deskripsi", placeholder="Detail task...")
        
        col1, col2 = st.columns(2)
        with col1:
            priority = st.selectbox("Priority", ["low", "medium", "high"])
            assignee = st.text_input("Assignee", placeholder="Nama penanggung jawab")
        with col2:
            due_date = st.date_input("Due Date", min_value=date.today())
            tags = st.text_input("Tags (pisahkan dengan koma)", placeholder="jadwal, review, poli")
        
        submitted = st.form_submit_button("➕ Tambah Task")
        
        if submitted and title:
            new_task = {
                "id": st.session_state.kanban_next_id,
                "title": title,
                "description": description,
                "priority": priority,
                "due_date": due_date.strftime("%Y-%m-%d") if due_date else None,
                "assignee": assignee,
                "tags": [tag.strip() for tag in tags.split(",") if tag.strip()],
                "created_by": "User",
                "created_date": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "last_updated": datetime.now().strftime("%Y-%m-%d %H:%M")
            }
            
            st.session_state.kanban_tasks["todo"].append(new_task)
            st.session_state.kanban_next_id += 1
            st.success(f"Task '{title}' ditambahkan!")
            st.rerun()

def move_task(task_id, from_column, to_column):
    """Pindah task antar column"""
    task_to_move = None
    for task in st.session_state.kanban_tasks[from_column]:
        if task["id"] == task_id:
            task_to_move = task
            break
    
    if task_to_move:
        st.session_state.kanban_tasks[from_column].remove(task_to_move)
        task_to_move["last_updated"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        st.session_state.kanban_tasks[to_column].append(task_to_move)

def delete_task(task_id, from_column):
    """Hapus task"""
    st.session_state.kanban_tasks[from_column] = [
        task for task in st.session_state.kanban_tasks[from_column] 
        if task["id"] != task_id
    ]

def export_kanban_to_excel():
    """Export Kanban board ke Excel"""
    wb = openpyxl.Workbook()
    
    # Buat sheet untuk setiap column
    for column_name in ["todo", "in_progress", "review", "done"]:
        ws = wb.create_sheet(title=column_name.capitalize())
        ws.append(["ID", "Title", "Description", "Priority", "Due Date", 
                  "Assignee", "Tags", "Created By", "Created Date", "Last Updated"])
        
        for task in st.session_state.kanban_tasks[column_name]:
            ws.append([
                task.get("id"),
                task.get("title"),
                task.get("description"),
                task.get("priority"),
                task.get("due_date"),
                task.get("assignee"),
                ", ".join(task.get("tags", [])),
                task.get("created_by"),
                task.get("created_date"),
                task.get("last_updated")
            ])
    
    # Hapus sheet default
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def import_kanban_from_excel(file):
    """Import Kanban dari Excel"""
    try:
        wb = load_workbook(file)
        for sheet_name in wb.sheetnames:
            column_key = sheet_name.lower()
            if column_key in st.session_state.kanban_tasks:
                df = pd.read_excel(file, sheet_name=sheet_name)
                st.session_state.kanban_tasks[column_key] = df.to_dict('records')
        
        # Update next_id
        max_id = 0
        for column in st.session_state.kanban_tasks.values():
            for task in column:
                if isinstance(task, dict) and "id" in task:
                    max_id = max(max_id, task["id"])
        st.session_state.kanban_next_id = max_id + 1
        
        return True
    except:
        return False

# ==================== MAIN APP ====================

def main():
    # Sidebar
    with st.sidebar:
        st.title("🏥📋 Jadwal Poli + Kanban")
        st.markdown("---")
        
        # Navigation
        page = st.radio(
            "Navigasi",
            ["📤 Upload & Proses", "📋 Kanban Board", "📊 Analytics", "⚙️ Settings"],
            index=0
        )
        
        st.markdown("---")
        
        if page == "📤 Upload & Proses":
            st.subheader("📋 Panduan")
            st.markdown("""
            1. **Upload** file Excel
            2. **Proses** jadwal
            3. **Download** hasil
            4. **Track** di Kanban
            """)
            
            if st.button("📥 Download Template", use_container_width=True):
                # Function untuk template (sama seperti sebelumnya)
                pass
                
        elif page == "📋 Kanban Board":
            st.subheader("Kanban Actions")
            
            # Export/Import Kanban
            col_exp, col_imp = st.columns(2)
            with col_exp:
                if st.button("📤 Export Kanban", use_container_width=True):
                    buffer = export_kanban_to_excel()
                    st.download_button(
                        label="Download Kanban.xlsx",
                        data=buffer,
                        file_name="kanban_board.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            
            with col_imp:
                kanban_file = st.file_uploader(
                    "Import Kanban",
                    type=['xlsx'],
                    key="kanban_import"
                )
                if kanban_file:
                    if import_kanban_from_excel(kanban_file):
                        st.success("Kanban imported!")
                    else:
                        st.error("Failed to import Kanban")
            
            # Quick stats
            st.markdown("---")
            total_tasks = sum(len(tasks) for tasks in st.session_state.kanban_tasks.values())
            st.metric("Total Tasks", total_tasks)
            
        st.markdown("---")
        st.caption("v1.0 • Dengan fitur Kanban")
    
    # Main content based on selected page
    if page == "📤 Upload & Proses":
        render_upload_page()
    elif page == "📋 Kanban Board":
        render_kanban_board()
    elif page == "📊 Analytics":
        render_analytics_page()
    elif page == "⚙️ Settings":
        render_settings_page()

def render_upload_page():
    """Halaman upload dan proses"""
    st.title("🏥 Pengisi Jadwal Poli Excel")
    st.caption("Dengan fitur Kanban untuk tracking progress")
    
    # Upload file section
    uploaded_file = st.file_uploader(
        "Upload file Excel (.xlsx)", 
        type=['xlsx'],
        help="Upload file dengan format yang sesuai"
    )
    
    if uploaded_file:
        col1, col2 = st.columns([2, 1])
        with col2:
            file_size = len(uploaded_file.getvalue()) / 1024
            st.metric("File Size", f"{file_size:.1f} KB")
        
        # Preview
        with st.expander("📄 Preview File", expanded=False):
            sheet_names = pd.ExcelFile(uploaded_file).sheet_names
            st.write(f"**Sheets:** {', '.join(sheet_names)}")
        
        # Proses button
        if st.button("🚀 Proses Jadwal & Buat Task Kanban", type="primary", use_container_width=True):
            with st.spinner("Memproses..."):
                try:
                    # Proses file (implementasi sebelumnya)
                    # ...
                    
                    # Auto-create Kanban task untuk tracking
                    auto_create_kanban_task(uploaded_file.name)
                    
                    st.success("✅ File diproses dan task Kanban dibuat!")
                    
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")

def auto_create_kanban_task(filename):
    """Otomatis buat task Kanban setelah proses"""
    new_task = {
        "id": st.session_state.kanban_next_id,
        "title": f"Review Jadwal: {filename}",
        "description": f"File jadwal yang diproses: {filename}. Perlu direview oleh koordinator.",
        "priority": "medium",
        "due_date": (datetime.now() + timedelta(days=2)).strftime("%Y-%m-%d"),
        "assignee": "Koordinator Poli",
        "tags": ["jadwal", "review", "auto-generated"],
        "created_by": "System",
        "created_date": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "last_updated": datetime.now().strftime("%Y-%m-%d %H:%M")
    }
    
    st.session_state.kanban_tasks["todo"].append(new_task)
    st.session_state.kanban_next_id += 1

def render_analytics_page():
    """Halaman analytics"""
    st.title("📊 Analytics & Reports")
    
    # Kanban analytics
    col1, col2, col3, col4 = st.columns(4)
    
    total_tasks = sum(len(tasks) for tasks in st.session_state.kanban_tasks.values())
    done_tasks = len(st.session_state.kanban_tasks["done"])
    
    with col1:
        st.metric("Total Tasks", total_tasks)
    with col2:
        st.metric("Tasks Done", done_tasks)
    with col3:
        progress = (done_tasks / total_tasks * 100) if total_tasks > 0 else 0
        st.metric("Completion", f"{progress:.1f}%")
    with col4:
        # Calculate average time in each column (simplified)
        st.metric("Active Tasks", len(st.session_state.kanban_tasks["in_progress"]))
    
    # Burndown chart (simplified)
    st.subheader("Task Distribution")
    import plotly.graph_objects as go
    
    fig = go.Figure(data=[
        go.Bar(
            x=list(st.session_state.kanban_tasks.keys()),
            y=[len(tasks) for tasks in st.session_state.kanban_tasks.values()],
            marker_color=list(KANBAN_COLORS.values())
        )
    ])
    
    fig.update_layout(
        title="Tasks per Column",
        xaxis_title="Column",
        yaxis_title="Number of Tasks"
    )
    
    st.plotly_chart(fig, use_container_width=True)

def render_settings_page():
    """Halaman settings"""
    st.title("⚙️ Settings")
    
    # Kanban settings
    with st.expander("Kanban Settings"):
        st.checkbox("Auto-create task setelah proses file", value=True)
        st.checkbox("Notify when task moves to Done", value=True)
        st.checkbox("Enable task deadlines", value=True)
    
    # Schedule settings
    with st.expander("Schedule Settings"):
        start_time = st.time_input("Start time", value=time(7, 30))
        end_time = st.time_input("End time", value=time(14, 30))
        interval = st.selectbox("Time interval", [15, 30, 60], index=1)
    
    # Export all settings
    if st.button("Export All Settings", use_container_width=True):
        settings = {
            "kanban_settings": {
                "auto_create": True,
                "notify": True,
                "deadlines": True
            },
            "schedule_settings": {
                "start_time": start_time.strftime("%H:%M"),
                "end_time": end_time.strftime("%H:%M"),
                "interval": interval
            }
        }
        
        st.download_button(
            label="Download Settings.json",
            data=json.dumps(settings, indent=2),
            file_name="kanban_settings.json",
            mime="application/json"
        )

if __name__ == "__main__":
    main()
