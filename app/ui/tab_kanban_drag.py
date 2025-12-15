# app/ui/tab_kanban_drag.py

import streamlit as st
import json
import pandas as pd
import numpy as np
from datetime import datetime

# ============================================================
# DEFAULT KANBAN UNTUK JADWAL DOKTER
# ============================================================
DEFAULT_KANBAN = {
    "⚠️ MASALAH JADWAL": [
        {"text": "Slot overload Poleks (>7)", "label": "Overload", "priority": "High"},
        {"text": "Dokter konflik waktu", "label": "Konflik", "priority": "High"},
        {"text": "Poli tidak terisi di jam sibuk", "label": "Kosong", "priority": "Medium"},
    ],
    "🔧 PERLU PENYESUAIAN": [
        {"text": "Distribusi tidak merata pagi-sore", "label": "Distribusi", "priority": "Medium"},
        {"text": "Dokter dengan jadwal terlalu padat", "label": "Beban", "priority": "Medium"},
    ],
    "⏳ DALAM PROSES": [
        {"text": "Review jadwal Poli Anak", "label": "Review", "priority": "Low"},
    ],
    "✅ OPTIMAL": [
        {"text": "Poli Jantung - distribusi bagus", "label": "Optimal", "priority": "Low"},
        {"text": "Jam 10:00-12:00 - slot ideal", "label": "Optimal", "priority": "Low"},
    ],
}

# ============================================================
# PRIORITY COLORS
# ============================================================
PRIORITY_COLORS = {
    "High": "#ff4d4f",    # Merah
    "Medium": "#faad14",   # Kuning
    "Low": "#52c41a"       # Hijau
}

LABEL_COLORS = {
    "Overload": "#ff7875",
    "Konflik": "#ff9c6e",
    "Kosong": "#69c0ff",
    "Distribusi": "#95de64",
    "Beban": "#b37feb",
    "Review": "#ffd666",
    "Optimal": "#5cdbd3"
}

# ============================================================
# SESSION MANAGEMENT
# ============================================================
def get_kanban_data():
    """Get kanban data from session state"""
    if "kanban_data" not in st.session_state:
        st.session_state["kanban_data"] = DEFAULT_KANBAN.copy()
    return st.session_state["kanban_data"]

def save_kanban_data(data):
    """Save kanban data to session state"""
    st.session_state["kanban_data"] = data

def get_schedule_issues():
    """Extract issues from processed schedule data"""
    issues = {
        "⚠️ MASALAH JADWAL": [],
        "🔧 PERLU PENYESUAIAN": [],
        "⏳ DALAM PROSES": [],
        "✅ OPTIMAL": []
    }
    
    if "processed_data" not in st.session_state:
        return issues
    
    df = st.session_state["processed_data"]
    slot_strings = st.session_state.get("slot_strings", [])
    
    if df is None or df.empty or not slot_strings:
        return issues
    
    # ============================================================
    # ANALYZE SCHEDULE FOR ISSUES
    # ============================================================
    
    # 1. Check for overload slots
    overload_issues = analyze_overload_slots(df, slot_strings)
    issues["⚠️ MASALAH JADWAL"].extend(overload_issues)
    
    # 2. Check for doctor conflicts
    conflict_issues = analyze_doctor_conflicts(df, slot_strings)
    issues["⚠️ MASALAH JADWAL"].extend(conflict_issues)
    
    # 3. Check for empty slots during peak hours
    empty_issues = analyze_empty_slots(df, slot_strings)
    issues["🔧 PERLU PENYESUAIAN"].extend(empty_issues)
    
    # 4. Check for distribution issues
    distribution_issues = analyze_distribution(df, slot_strings)
    issues["🔧 PERLU PENYESUAIAN"].extend(distribution_issues)
    
    # 5. Find optimal schedules
    optimal_issues = find_optimal_schedules(df, slot_strings)
    issues["✅ OPTIMAL"].extend(optimal_issues)
    
    return issues

def analyze_overload_slots(df, slot_strings):
    """Analyze slots with too many Poleks"""
    issues = []
    max_poleks = st.session_state.get("config", type('obj', (object,), {'max_poleks_per_slot': 7})).max_poleks_per_slot
    
    for hari in df["HARI"].unique():
        hari_data = df[df["HARI"] == hari]
        
        for slot in slot_strings[:15]:  # Check first 15 slots
            if slot in hari_data.columns:
                poleks_count = (hari_data[slot] == "E").sum()
                
                if poleks_count > max_poleks:
                    issues.append({
                        "text": f"{hari} {slot}: {poleks_count} Poleks (batas {max_poleks})",
                        "label": "Overload",
                        "priority": "High",
                        "data": {
                            "hari": hari,
                            "slot": slot,
                            "count": int(poleks_count),  # Convert to Python int
                            "max": max_poleks,
                            "type": "overload"
                        }
                    })
    
    return issues

def analyze_doctor_conflicts(df, slot_strings):
    """Analyze doctors with schedule conflicts"""
    issues = []
    
    for (dokter, hari), group in df.groupby(["DOKTER", "HARI"]):
        if len(group) > 1:  # Doctor appears in more than 1 poli
            conflict_slots = []
            
            for slot in slot_strings[:10]:  # Check first 10 slots
                if slot in group.columns:
                    active_polis = group[group[slot].isin(["R", "E"])]["POLI"].tolist()
                    if len(active_polis) > 1:
                        conflict_slots.append({
                            "slot": slot,
                            "polis": active_polis
                        })
            
            if conflict_slots:
                issues.append({
                    "text": f"Dr. {dokter} - {hari}: konflik di {len(conflict_slots)} slot",
                    "label": "Konflik",
                    "priority": "High",
                    "data": {
                        "dokter": dokter,
                        "hari": hari,
                        "conflicts": conflict_slots,
                        "type": "conflict"
                    }
                })
    
    return issues

def analyze_empty_slots(df, slot_strings):
    """Analyze empty slots during peak hours"""
    issues = []
    
    # Define peak hours (10:00-12:00)
    peak_slots = [s for s in slot_strings if "10:00" <= s <= "12:00"]
    
    for hari in df["HARI"].unique():
        hari_data = df[df["HARI"] == hari]
        
        for poli in hari_data["POLI"].unique():
            poli_data = hari_data[hari_data["POLI"] == poli]
            
            empty_in_peak = 0
            for slot in peak_slots:
                if slot in poli_data.columns:
                    # Check if slot has any R or E values
                    slot_values = poli_data[slot].values
                    if not any(val in ["R", "E"] for val in slot_values):
                        empty_in_peak += 1
            
            if empty_in_peak >= 2:  # At least 2 empty slots in peak hours
                issues.append({
                    "text": f"{poli} - {hari}: {empty_in_peak} slot kosong di jam sibuk",
                    "label": "Kosong",
                    "priority": "Medium",
                    "data": {
                        "poli": poli,
                        "hari": hari,
                        "empty_slots": int(empty_in_peak),  # Convert to Python int
                        "type": "empty"
                    }
                })
    
    return issues

def analyze_distribution(df, slot_strings):
    """Analyze distribution issues"""
    issues = []
    
    for poli in df["POLI"].unique():
        poli_data = df[df["POLI"] == poli]
        
        # Count morning vs afternoon slots
        morning_slots = [s for s in slot_strings if s < "12:00"]
        afternoon_slots = [s for s in slot_strings if s >= "12:00"]
        
        # Initialize counts
        morning_count = 0
        afternoon_count = 0
        
        # Count R and E slots for morning
        for slot in morning_slots:
            if slot in poli_data.columns:
                morning_count += int((poli_data[slot] == "R").sum())
                morning_count += int((poli_data[slot] == "E").sum())
        
        # Count R and E slots for afternoon
        for slot in afternoon_slots:
            if slot in poli_data.columns:
                afternoon_count += int((poli_data[slot] == "R").sum())
                afternoon_count += int((poli_data[slot] == "E").sum())
        
        total = morning_count + afternoon_count
        if total > 0:
            morning_pct = (morning_count / total) * 100
            afternoon_pct = (afternoon_count / total) * 100
            
            if morning_pct > 70:  # More than 70% in morning
                issues.append({
                    "text": f"{poli}: {morning_pct:.0f}% pagi, {afternoon_pct:.0f}% sore",
                    "label": "Distribusi",
                    "priority": "Medium",
                    "data": {
                        "poli": poli,
                        "morning_pct": float(morning_pct),  # Convert to float
                        "afternoon_pct": float(afternoon_pct),  # Convert to float
                        "type": "distribution"
                    }
                })
    
    return issues

def find_optimal_schedules(df, slot_strings):
    """Find optimal schedules to highlight"""
    issues = []
    
    # Find slots with good distribution (3-5 doctors)
    for hari in df["HARI"].unique():
        hari_data = df[df["HARI"] == hari]
        
        for slot in ["10:00", "11:00", "13:00"]:  # Key time slots
            if slot in hari_data.columns:
                doctor_count = (hari_data[slot].isin(["R", "E"])).sum()
                
                if 3 <= doctor_count <= 5:  # Optimal range
                    issues.append({
                        "text": f"{hari} {slot}: {int(doctor_count)} dokter (optimal)",  # Convert to int
                        "label": "Optimal",
                        "priority": "Low",
                        "data": {
                            "hari": hari,
                            "slot": slot,
                            "doctor_count": int(doctor_count),  # Convert to Python int
                            "type": "optimal"
                        }
                    })
    
    return issues

# Custom JSON encoder untuk handle numpy types
class NumpyJSONEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, np.integer):
            return int(obj)
        elif isinstance(obj, np.floating):
            return float(obj)
        elif isinstance(obj, np.ndarray):
            return obj.tolist()
        elif isinstance(obj, pd.Timestamp):
            return obj.isoformat()
        return super().default(obj)

# ============================================================
# RENDER TAB KANBAN
# ============================================================
def render_drag_kanban():
    st.title("📌 Kanban Board - Manajemen Jadwal Dokter")
    
    # Get current kanban data
    kanban_data = get_kanban_data()
    
    # Sidebar controls
    with st.sidebar:
        st.header("⚙️ Kontrol Kanban")
        
        # Generate cards from schedule
        if st.button("🔄 Generate dari Jadwal", width='stretch'):
            if "processed_data" in st.session_state:
                issues = get_schedule_issues()
                
                # Clear existing data except "DALAM PROSES" and "OPTIMAL"
                kanban_data["⚠️ MASALAH JADWAL"] = issues["⚠️ MASALAH JADWAL"]
                kanban_data["🔧 PERLU PENYESUAIAN"] = issues["🔧 PERLU PENYESUAIAN"]
                
                # Keep "DALAM PROSES" and "OPTIMAL" as they are
                save_kanban_data(kanban_data)
                st.success(f"Generated {len(issues['⚠️ MASALAH JADWAL'])} masalah dan {len(issues['🔧 PERLU PENYESUAIAN'])} penyesuaian")
                st.rerun()
            else:
                st.warning("Belum ada data jadwal yang diproses")
        
        st.divider()
        
        # Add new card
        st.subheader("➕ Tambah Kartu Manual")
        
        with st.form("add_card_form"):
            new_text = st.text_input("Judul Kartu")
            new_label = st.selectbox("Label", list(LABEL_COLORS.keys()))
            new_priority = st.selectbox("Prioritas", list(PRIORITY_COLORS.keys()))
            target_column = st.selectbox("Kolom Tujuan", list(kanban_data.keys()))
            
            if st.form_submit_button("Tambah Kartu", width='stretch'):
                if new_text.strip():
                    kanban_data[target_column].append({
                        "text": new_text,
                        "label": new_label,
                        "priority": new_priority
                    })
                    save_kanban_data(kanban_data)
                    st.success("Kartu ditambahkan!")
                    st.rerun()
        
        st.divider()
        
        # Import/Export
        st.subheader("📁 Import/Export")
        
        col1, col2 = st.columns(2)
        with col1:
            # Create download button
            json_str = json.dumps(kanban_data, indent=2, ensure_ascii=False, cls=NumpyJSONEncoder)
            st.download_button(
                label="📥 Download JSON",
                data=json_str,
                file_name=f"kanban_jadwal_{datetime.now().strftime('%Y%m%d')}.json",
                mime="application/json",
                use_container_width=True
            )
        
        with col2:
            uploaded = st.file_uploader("Upload JSON", type=["json"], label_visibility="collapsed")
            if uploaded:
                try:
                    loaded = json.load(uploaded)
                    save_kanban_data(loaded)
                    st.success("JSON berhasil diimpor!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")
        
        st.divider()
        
        # Reset buttons
        if st.button("🗑️ Reset ke Default", width='stretch'):
            save_kanban_data(DEFAULT_KANBAN.copy())
            st.success("Reset berhasil!")
            st.rerun()
        
        if st.button("🧹 Kosongkan Semua", width='stretch', type="secondary"):
            for column in kanban_data:
                kanban_data[column] = []
            save_kanban_data(kanban_data)
            st.success("Semua kartu dihapus!")
            st.rerun()
    
    # Main kanban board
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown("### ⚠️ MASALAH JADWAL")
        st.caption(f"{len(kanban_data['⚠️ MASALAH JADWAL'])} kartu")
        render_column("⚠️ MASALAH JADWAL", kanban_data)
    
    with col2:
        st.markdown("### 🔧 PERLU PENYESUAIAN")
        st.caption(f"{len(kanban_data['🔧 PERLU PENYESUAIAN'])} kartu")
        render_column("🔧 PERLU PENYESUAIAN", kanban_data)
    
    with col3:
        st.markdown("### ⏳ DALAM PROSES")
        st.caption(f"{len(kanban_data['⏳ DALAM PROSES'])} kartu")
        render_column("⏳ DALAM PROSES", kanban_data)
    
    with col4:
        st.markdown("### ✅ OPTIMAL")
        st.caption(f"{len(kanban_data['✅ OPTIMAL'])} kartu")
        render_column("✅ OPTIMAL", kanban_data)
    
    # Interactive HTML Kanban Board
    st.divider()
    st.subheader("🎯 Drag & Drop Board")
    
    # Prepare data for HTML
    html_data = json.dumps(kanban_data, ensure_ascii=False, cls=NumpyJSONEncoder)
    
    # Generate HTML with interactive kanban
    html = generate_kanban_html(html_data)
    st.components.v1.html(html, height=700, scrolling=True)
    
    # Statistics
    st.divider()
    st.subheader("📊 Statistik Kanban")
    
    total_cards = sum(len(cards) for cards in kanban_data.values())
    high_priority = sum(1 for column in kanban_data.values() 
                       for card in column if card.get("priority") == "High")
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Kartu", total_cards)
    col2.metric("Prioritas Tinggi", high_priority)
    col3.metric("Masalah", len(kanban_data["⚠️ MASALAH JADWAL"]))
    col4.metric("Optimal", len(kanban_data["✅ OPTIMAL"]))

def render_column(column_name, kanban_data):
    """Render a single column in Streamlit"""
    cards = kanban_data[column_name]
    
    if not cards:
        st.info("Tidak ada kartu")
        return
    
    for i, card in enumerate(cards):
        with st.container():
            # Priority indicator
            priority_color = PRIORITY_COLORS.get(card.get("priority", "Medium"), "#d9d9d9")
            
            # Card header with priority
            st.markdown(f"""
            <div style="border-left: 4px solid {priority_color}; padding-left: 10px; margin-bottom: 10px;">
                <div style="font-weight: 500;">{card['text']}</div>
                <div style="display: flex; gap: 8px; margin-top: 4px;">
                    <span style="background-color: {LABEL_COLORS.get(card['label'], '#d9d9d9')}; 
                           color: white; padding: 2px 8px; border-radius: 12px; font-size: 0.8em;">
                        {card['label']}
                    </span>
                    <span style="color: {priority_color}; font-size: 0.8em;">
                        {card.get('priority', 'Medium')}
                    </span>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            # Action buttons
            col1, col2, col3 = st.columns([1, 1, 2])
            with col1:
                if st.button("📋", key=f"copy_{column_name}_{i}", help="Salin"):
                    st.toast(f"Disalin: {card['text']}")
            
            with col2:
                if st.button("✏️", key=f"edit_{column_name}_{i}", help="Edit"):
                    # Simple edit in place
                    new_text = st.text_input("Edit teks:", value=card['text'], 
                                           key=f"edit_input_{column_name}_{i}")
                    if new_text != card['text']:
                        kanban_data[column_name][i]["text"] = new_text
                        save_kanban_data(kanban_data)
                        st.rerun()
            
            with col3:
                if st.button("🗑️", key=f"delete_{column_name}_{i}", help="Hapus"):
                    kanban_data[column_name].pop(i)
                    save_kanban_data(kanban_data)
                    st.rerun()
            
            st.divider()

def generate_kanban_html(kanban_data_json):
    """Generate interactive HTML kanban board"""
    return f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Kanban Jadwal Dokter</title>
    <style>
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            padding: 20px;
            background: #f8f9fa;
        }}
        .kanban-header {{
            text-align: center;
            margin-bottom: 30px;
            color: #2c3e50;
        }}
        .board {{
            display: flex;
            gap: 20px;
            overflow-x: auto;
            padding: 20px;
        }}
        .column {{
            background: white;
            border-radius: 12px;
            padding: 20px;
            min-width: 280px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }}
        .column-header {{
            font-weight: 600;
            font-size: 16px;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid #e0e0e0;
        }}
        .card-list {{
            min-height: 400px;
            padding: 10px;
            background: #f8f9fa;
            border-radius: 8px;
        }}
        .card {{
            background: white;
            padding: 15px;
            margin-bottom: 12px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            cursor: move;
            border-left: 4px solid #4CAF50;
            transition: transform 0.2s;
        }}
        .card:hover {{
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        }}
        .card-label {{
            display: inline-block;
            padding: 3px 10px;
            border-radius: 12px;
            font-size: 12px;
            font-weight: 500;
            margin-bottom: 8px;
            color: white;
        }}
        .card-content {{
            font-size: 14px;
            line-height: 1.4;
        }}
        .card-priority {{
            font-size: 11px;
            margin-top: 8px;
            padding: 2px 8px;
            border-radius: 10px;
            display: inline-block;
        }}
        .priority-high {{ background: #ff4d4f; color: white; }}
        .priority-medium {{ background: #faad14; color: white; }}
        .priority-low {{ background: #52c41a; color: white; }}
        
        .label-overload {{ background: #ff7875; }}
        .label-konflik {{ background: #ff9c6e; }}
        .label-kosong {{ background: #69c0ff; }}
        .label-distribusi {{ background: #95de64; }}
        .label-beban {{ background: #b37feb; }}
        .label-review {{ background: #ffd666; color: #333; }}
        .label-optimal {{ background: #5cdbd3; }}
        
        .controls {{
            margin-bottom: 20px;
            text-align: center;
        }}
        .export-btn {{
            background: #1890ff;
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 6px;
            cursor: pointer;
            font-size: 14px;
        }}
        .export-btn:hover {{
            background: #40a9ff;
        }}
    </style>
    <script src="https://cdn.jsdelivr.net/npm/sortablejs@1.15.0/Sortable.min.js"></script>
</head>
<body>
    <div class="kanban-header">
        <h2>🎯 Kanban Board - Manajemen Jadwal Dokter</h2>
        <p>Drag & drop untuk mengatur prioritas penjadwalan</p>
    </div>
    
    <div class="controls">
        <button class="export-btn" onclick="exportKanban()">📥 Export Perubahan</button>
        <span id="save-status" style="margin-left: 20px; color: #52c41a;"></span>
    </div>
    
    <div class="board" id="kanban-board"></div>
    
    <script>
        // Initial kanban data
        const kanbanData = {kanban_data_json};
        
        // Color mappings
        const labelColors = {json.dumps(LABEL_COLORS, cls=NumpyJSONEncoder)};
        const priorityColors = {json.dumps(PRIORITY_COLORS, cls=NumpyJSONEncoder)};
        
        // Render the board
        function renderBoard() {{
            const board = document.getElementById('kanban-board');
            board.innerHTML = '';
            
            Object.entries(kanbanData).forEach(([columnName, cards]) => {{
                const columnDiv = document.createElement('div');
                columnDiv.className = 'column';
                
                const header = document.createElement('div');
                header.className = 'column-header';
                header.textContent = columnName + ` (${{cards.length}})`;
                
                const cardList = document.createElement('div');
                cardList.className = 'card-list';
                cardList.dataset.column = columnName;
                
                cards.forEach(card => {{
                    const cardDiv = document.createElement('div');
                    cardDiv.className = 'card';
                    cardDiv.dataset.card = JSON.stringify(card);
                    
                    const label = document.createElement('div');
                    label.className = `card-label label-${{card.label.toLowerCase()}}`;
                    label.textContent = card.label;
                    label.style.backgroundColor = labelColors[card.label] || '#d9d9d9';
                    
                    const content = document.createElement('div');
                    content.className = 'card-content';
                    content.textContent = card.text;
                    
                    const priority = document.createElement('div');
                    priority.className = `card-priority priority-${{card.priority.toLowerCase()}}`;
                    priority.textContent = card.priority;
                    priority.style.backgroundColor = priorityColors[card.priority] || '#d9d9d9';
                    
                    cardDiv.appendChild(label);
                    cardDiv.appendChild(content);
                    cardDiv.appendChild(priority);
                    cardList.appendChild(cardDiv);
                }});
                
                columnDiv.appendChild(header);
                columnDiv.appendChild(cardList);
                board.appendChild(columnDiv);
                
                // Make list sortable
                new Sortable(cardList, {{
                    group: 'shared',
                    animation: 150,
                    ghostClass: 'sortable-ghost',
                    chosenClass: 'sortable-chosen',
                    dragClass: 'sortable-drag',
                    onEnd: function(evt) {{
                        updateKanbanData();
                        showSaveStatus('Perubahan disimpan');
                    }}
                }});
            }});
        }}
        
        // Update kanban data after drag & drop
        function updateKanbanData() {{
            const columns = document.querySelectorAll('.card-list');
            
            columns.forEach(columnElement => {{
                const columnName = columnElement.dataset.column;
                const cards = Array.from(columnElement.children).map(cardElement => {{
                    return JSON.parse(cardElement.dataset.card);
                }});
                
                kanbanData[columnName] = cards;
            }});
        }}
        
        // Export kanban data
        function exportKanban() {{
            updateKanbanData();
            
            const dataStr = JSON.stringify(kanbanData, null, 2);
            const dataBlob = new Blob([dataStr], {{ type: 'application/json' }});
            
            const url = URL.createObjectURL(dataBlob);
            const link = document.createElement('a');
            link.href = url;
            link.download = 'kanban_jadwal_' + new Date().toISOString().split('T')[0] + '.json';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            
            showSaveStatus('File berhasil di-download!');
        }}
        
        // Show save status
        function showSaveStatus(message) {{
            const statusEl = document.getElementById('save-status');
            statusEl.textContent = message;
            setTimeout(() => {{
                statusEl.textContent = '';
            }}, 3000);
        }}
        
        // Initialize
        renderBoard();
    </script>
</body>
</html>
"""

# For backward compatibility
if __name__ == "__main__":
    render_drag_kanban()
