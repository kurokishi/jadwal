# app/ui/tab_kanban_drag.py

import streamlit as st
import json
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import plotly.graph_objects as go
import plotly.express as px

# ============================================================
# DEFAULT KANBAN UNTUK JADWAL DOKTER
# ============================================================
DEFAULT_KANBAN = {
    "⚠️ MASALAH JADWAL": [
        {"id": "1", "text": "Slot overload Poleks (>7)", "label": "Overload", "priority": "High", 
         "created": "2024-01-15", "due_date": "2024-01-20", "assignee": "Admin"},
        {"id": "2", "text": "Dokter konflik waktu", "label": "Konflik", "priority": "High",
         "created": "2024-01-15", "due_date": "2024-01-18", "assignee": "Manager"},
        {"id": "3", "text": "Poli tidak terisi di jam sibuk", "label": "Kosong", "priority": "Medium",
         "created": "2024-01-14", "due_date": "2024-01-22", "assignee": "Admin"},
    ],
    "🔧 PERLU PENYESUAIAN": [
        {"id": "4", "text": "Distribusi tidak merata pagi-sore", "label": "Distribusi", "priority": "Medium",
         "created": "2024-01-13", "due_date": "2024-01-25", "assignee": "Staff"},
        {"id": "5", "text": "Dokter dengan jadwal terlalu padat", "label": "Beban", "priority": "Medium",
         "created": "2024-01-12", "due_date": "2024-01-19", "assignee": "Manager"},
    ],
    "⏳ DALAM PROSES": [
        {"id": "6", "text": "Review jadwal Poli Anak", "label": "Review", "priority": "Low",
         "created": "2024-01-10", "due_date": "2024-01-30", "assignee": "Supervisor"},
    ],
    "✅ OPTIMAL": [
        {"id": "7", "text": "Poli Jantung - distribusi bagus", "label": "Optimal", "priority": "Low",
         "created": "2024-01-05", "due_date": "", "assignee": "System"},
        {"id": "8", "text": "Jam 10:00-12:00 - slot ideal", "label": "Optimal", "priority": "Low",
         "created": "2024-01-06", "due_date": "", "assignee": "System"},
    ],
}

# ============================================================
# COLORS & THEMES
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

COLUMN_COLORS = {
    "⚠️ MASALAH JADWAL": "#fff2f0",
    "🔧 PERLU PENYESUAIAN": "#fffbe6",
    "⏳ DALAM PROSES": "#f0f5ff",
    "✅ OPTIMAL": "#f6ffed"
}

COLUMN_BORDER_COLORS = {
    "⚠️ MASALAH JADWAL": "#ffccc7",
    "🔧 PERLU PENYESUAIAN": "#ffe58f",
    "⏳ DALAM PROSES": "#adc6ff",
    "✅ OPTIMAL": "#b7eb8f"
}

# ============================================================
# SESSION MANAGEMENT
# ============================================================
def get_kanban_data():
    """Get kanban data from session state"""
    if "kanban_data" not in st.session_state:
        st.session_state["kanban_data"] = DEFAULT_KANBAN.copy()
    
    # Ensure all cards have IDs
    for column_name, cards in st.session_state["kanban_data"].items():
        for i, card in enumerate(cards):
            if "id" not in card:
                card["id"] = f"{column_name[:2]}_{i}_{datetime.now().timestamp()}"
            if "created" not in card:
                card["created"] = datetime.now().strftime("%Y-%m-%d")
            if "due_date" not in card:
                card["due_date"] = ""
            if "assignee" not in card:
                card["assignee"] = "Unassigned"
    
    return st.session_state["kanban_data"]

def save_kanban_data(data):
    """Save kanban data to session state"""
    st.session_state["kanban_data"] = data
    # Save to browser storage via session state
    st.session_state["last_saved"] = datetime.now().strftime("%H:%M:%S")

def get_next_card_id():
    """Generate next card ID"""
    kanban_data = get_kanban_data()
    all_ids = []
    for column in kanban_data.values():
        for card in column:
            if "id" in card:
                all_ids.append(card["id"])
    
    # Find next numeric ID
    numeric_ids = []
    for card_id in all_ids:
        try:
            if "_" in card_id:
                numeric_ids.append(int(card_id.split("_")[1]))
        except:
            pass
    
    next_id = max(numeric_ids) + 1 if numeric_ids else 1
    return f"card_{next_id}"

# ============================================================
# SCHEDULE ISSUE ANALYSIS
# ============================================================
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
    
    # Analyze schedule for issues
    card_counter = len(get_all_cards()) + 1
    
    # 1. Check for overload slots
    overload_issues = analyze_overload_slots(df, slot_strings, card_counter)
    issues["⚠️ MASALAH JADWAL"].extend(overload_issues)
    card_counter += len(overload_issues)
    
    # 2. Check for doctor conflicts
    conflict_issues = analyze_doctor_conflicts(df, slot_strings, card_counter)
    issues["⚠️ MASALAH JADWAL"].extend(conflict_issues)
    card_counter += len(conflict_issues)
    
    # 3. Check for empty slots during peak hours
    empty_issues = analyze_empty_slots(df, slot_strings, card_counter)
    issues["🔧 PERLU PENYESUAIAN"].extend(empty_issues)
    card_counter += len(empty_issues)
    
    # 4. Check for distribution issues
    distribution_issues = analyze_distribution(df, slot_strings, card_counter)
    issues["🔧 PERLU PENYESUAIAN"].extend(distribution_issues)
    card_counter += len(distribution_issues)
    
    # 5. Find optimal schedules
    optimal_issues = find_optimal_schedules(df, slot_strings, card_counter)
    issues["✅ OPTIMAL"].extend(optimal_issues)
    
    return issues

def analyze_overload_slots(df, slot_strings, start_id):
    """Analyze slots with too many Poleks"""
    issues = []
    max_poleks = st.session_state.get("config", type('obj', (object,), {'max_poleks_per_slot': 7})).max_poleks_per_slot
    
    for hari in df["HARI"].unique():
        hari_data = df[df["HARI"] == hari]
        
        for slot in slot_strings[:15]:
            if slot in hari_data.columns:
                poleks_count = (hari_data[slot] == "E").sum()
                
                if poleks_count > max_poleks:
                    issues.append({
                        "id": f"overload_{start_id + len(issues)}",
                        "text": f"{hari} {slot}: {int(poleks_count)} Poleks (batas {max_poleks})",
                        "label": "Overload",
                        "priority": "High",
                        "created": datetime.now().strftime("%Y-%m-%d"),
                        "due_date": (datetime.now() + timedelta(days=3)).strftime("%Y-%m-%d"),
                        "assignee": "Admin",
                        "data": {
                            "hari": hari,
                            "slot": slot,
                            "count": int(poleks_count),
                            "max": max_poleks,
                            "type": "overload"
                        }
                    })
    
    return issues

def analyze_doctor_conflicts(df, slot_strings, start_id):
    """Analyze doctors with schedule conflicts"""
    issues = []
    
    for (dokter, hari), group in df.groupby(["DOKTER", "HARI"]):
        if len(group) > 1:
            conflict_slots = []
            
            for slot in slot_strings[:10]:
                if slot in group.columns:
                    active_polis = group[group[slot].isin(["R", "E"])]["POLI"].tolist()
                    if len(active_polis) > 1:
                        conflict_slots.append({
                            "slot": slot,
                            "polis": active_polis
                        })
            
            if conflict_slots:
                issues.append({
                    "id": f"conflict_{start_id + len(issues)}",
                    "text": f"Dr. {dokter} - {hari}: konflik di {len(conflict_slots)} slot",
                    "label": "Konflik",
                    "priority": "High",
                    "created": datetime.now().strftime("%Y-%m-%d"),
                    "due_date": (datetime.now() + timedelta(days=2)).strftime("%Y-%m-%d"),
                    "assignee": "Manager",
                    "data": {
                        "dokter": dokter,
                        "hari": hari,
                        "conflicts": conflict_slots,
                        "type": "conflict"
                    }
                })
    
    return issues

def analyze_empty_slots(df, slot_strings, start_id):
    """Analyze empty slots during peak hours"""
    issues = []
    
    peak_slots = [s for s in slot_strings if "10:00" <= s <= "12:00"]
    
    for hari in df["HARI"].unique():
        hari_data = df[df["HARI"] == hari]
        
        for poli in hari_data["POLI"].unique():
            poli_data = hari_data[hari_data["POLI"] == poli]
            
            empty_in_peak = 0
            for slot in peak_slots:
                if slot in poli_data.columns:
                    slot_values = poli_data[slot].values
                    if not any(val in ["R", "E"] for val in slot_values):
                        empty_in_peak += 1
            
            if empty_in_peak >= 2:
                issues.append({
                    "id": f"empty_{start_id + len(issues)}",
                    "text": f"{poli} - {hari}: {empty_in_peak} slot kosong di jam sibuk",
                    "label": "Kosong",
                    "priority": "Medium",
                    "created": datetime.now().strftime("%Y-%m-%d"),
                    "due_date": (datetime.now() + timedelta(days=5)).strftime("%Y-%m-%d"),
                    "assignee": "Staff",
                    "data": {
                        "poli": poli,
                        "hari": hari,
                        "empty_slots": int(empty_in_peak),
                        "type": "empty"
                    }
                })
    
    return issues

def analyze_distribution(df, slot_strings, start_id):
    """Analyze distribution issues"""
    issues = []
    
    for poli in df["POLI"].unique():
        poli_data = df[df["POLI"] == poli]
        
        morning_slots = [s for s in slot_strings if s < "12:00"]
        afternoon_slots = [s for s in slot_strings if s >= "12:00"]
        
        morning_count = 0
        afternoon_count = 0
        
        for slot in morning_slots:
            if slot in poli_data.columns:
                morning_count += int((poli_data[slot] == "R").sum())
                morning_count += int((poli_data[slot] == "E").sum())
        
        for slot in afternoon_slots:
            if slot in poli_data.columns:
                afternoon_count += int((poli_data[slot] == "R").sum())
                afternoon_count += int((poli_data[slot] == "E").sum())
        
        total = morning_count + afternoon_count
        if total > 0:
            morning_pct = (morning_count / total) * 100
            afternoon_pct = (afternoon_count / total) * 100
            
            if morning_pct > 70:
                issues.append({
                    "id": f"dist_{start_id + len(issues)}",
                    "text": f"{poli}: {morning_pct:.0f}% pagi, {afternoon_pct:.0f}% sore",
                    "label": "Distribusi",
                    "priority": "Medium",
                    "created": datetime.now().strftime("%Y-%m-%d"),
                    "due_date": (datetime.now() + timedelta(days=7)).strftime("%Y-%m-%d"),
                    "assignee": "Analyst",
                    "data": {
                        "poli": poli,
                        "morning_pct": float(morning_pct),
                        "afternoon_pct": float(afternoon_pct),
                        "type": "distribution"
                    }
                })
    
    return issues

def find_optimal_schedules(df, slot_strings, start_id):
    """Find optimal schedules to highlight"""
    issues = []
    
    for hari in df["HARI"].unique():
        hari_data = df[df["HARI"] == hari]
        
        for slot in ["10:00", "11:00", "13:00"]:
            if slot in hari_data.columns:
                doctor_count = (hari_data[slot].isin(["R", "E"])).sum()
                
                if 3 <= doctor_count <= 5:
                    issues.append({
                        "id": f"opt_{start_id + len(issues)}",
                        "text": f"{hari} {slot}: {int(doctor_count)} dokter (optimal)",
                        "label": "Optimal",
                        "priority": "Low",
                        "created": datetime.now().strftime("%Y-%m-%d"),
                        "due_date": "",
                        "assignee": "System",
                        "data": {
                            "hari": hari,
                            "slot": slot,
                            "doctor_count": int(doctor_count),
                            "type": "optimal"
                        }
                    })
    
    return issues

# ============================================================
# HELPER FUNCTIONS
# ============================================================
def get_all_cards():
    """Get all cards from all columns"""
    kanban_data = get_kanban_data()
    all_cards = []
    for column in kanban_data.values():
        all_cards.extend(column)
    return all_cards

def get_card_statistics():
    """Calculate statistics for all cards"""
    kanban_data = get_kanban_data()
    
    stats = {
        "total_cards": 0,
        "high_priority": 0,
        "medium_priority": 0,
        "low_priority": 0,
        "overdue": 0,
        "due_today": 0,
        "by_assignee": {},
        "by_label": {},
        "completion_rate": 0
    }
    
    all_cards = get_all_cards()
    stats["total_cards"] = len(all_cards)
    
    today = datetime.now().date()
    
    for card in all_cards:
        # Priority count
        priority = card.get("priority", "Medium")
        if priority == "High":
            stats["high_priority"] += 1
        elif priority == "Medium":
            stats["medium_priority"] += 1
        else:
            stats["low_priority"] += 1
        
        # Assignee count
        assignee = card.get("assignee", "Unassigned")
        stats["by_assignee"][assignee] = stats["by_assignee"].get(assignee, 0) + 1
        
        # Label count
        label = card.get("label", "Unknown")
        stats["by_label"][label] = stats["by_label"].get(label, 0) + 1
        
        # Due date check
        due_date = card.get("due_date", "")
        if due_date:
            try:
                due = datetime.strptime(due_date, "%Y-%m-%d").date()
                if due < today:
                    stats["overdue"] += 1
                elif due == today:
                    stats["due_today"] += 1
            except:
                pass
    
    # Completion rate (cards in OPTIMAL vs total)
    optimal_cards = len(kanban_data.get("✅ OPTIMAL", []))
    if stats["total_cards"] > 0:
        stats["completion_rate"] = (optimal_cards / stats["total_cards"]) * 100
    
    return stats

def get_burndown_data():
    """Generate mock burndown data"""
    days = 7
    dates = [(datetime.now() - timedelta(days=i)).strftime("%Y-%m-%d") for i in range(days-1, -1, -1)]
    
    # Mock data - in real app, this would come from history
    issues = [12, 11, 10, 8, 7, 6, 5]
    completed = [0, 1, 2, 4, 5, 6, 7]
    
    return pd.DataFrame({
        "Tanggal": dates,
        "Masalah Aktif": issues,
        "Terselesaikan": completed
    })

# ============================================================
# CUSTOM JSON ENCODER
# ============================================================
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
        elif isinstance(obj, datetime):
            return obj.isoformat()
        return super().default(obj)

# ============================================================
# RENDER FUNCTIONS
# ============================================================
def render_card(card, column_name, index):
    """Render a single card with enhanced design"""
    priority_color = PRIORITY_COLORS.get(card.get("priority", "Medium"), "#d9d9d9")
    label_color = LABEL_COLORS.get(card.get("label", "Unknown"), "#d9d9d9")
    
    # Calculate due date status
    due_status = ""
    due_class = ""
    due_date = card.get("due_date", "")
    
    if due_date:
        try:
            due = datetime.strptime(due_date, "%Y-%m-%d").date()
            today = datetime.now().date()
            
            if due < today:
                due_status = "⏰ Terlambat"
                due_class = "overdue"
            elif due == today:
                due_status = "⏰ Hari ini"
                due_class = "due-today"
            elif (due - today).days <= 3:
                due_status = f"⏰ {(due - today).days} hari"
                due_class = "due-soon"
        except:
            pass
    
    # Card HTML
    card_html = f"""
    <div class="kanban-card" style="border-left-color: {priority_color};">
        <div class="card-header">
            <div class="card-text">{card['text']}</div>
            <div class="card-meta">
                <span class="card-label" style="background: {label_color};">{card.get('label', '')}</span>
                <span class="card-priority" style="color: {priority_color};">{card.get('priority', '')}</span>
            </div>
        </div>
        <div class="card-footer">
            <div class="card-assignee">👤 {card.get('assignee', 'Unassigned')}</div>
            <div class="card-dates">
                <small>📅 {card.get('created', '')}</small>
    """
    
    if due_status:
        card_html += f'<small class="{due_class}">{due_status}</small>'
    
    card_html += """
            </div>
        </div>
    </div>
    """
    
    # Render card with actions
    with st.container():
        # Card content
        st.markdown(card_html, unsafe_allow_html=True)
        
        # Action buttons in columns
        col1, col2, col3, col4 = st.columns([1, 1, 1, 2])
        
        with col1:
            if st.button("📋", key=f"copy_{card['id']}_{index}", help="Salin"):
                st.toast(f"Disalin: {card['text'][:50]}...")
        
        with col2:
            # Quick move to In Progress
            if column_name != "⏳ DALAM PROSES":
                if st.button("▶️", key=f"start_{card['id']}_{index}", help="Mulai proses"):
                    kanban_data = get_kanban_data()
                    kanban_data["⏳ DALAM PROSES"].append(card)
                    kanban_data[column_name] = [c for c in kanban_data[column_name] if c["id"] != card["id"]]
                    save_kanban_data(kanban_data)
                    st.rerun()
        
        with col3:
            # Quick move to Optimal
            if column_name != "✅ OPTIMAL":
                if st.button("✅", key=f"complete_{card['id']}_{index}", help="Tandai selesai"):
                    kanban_data = get_kanban_data()
                    kanban_data["✅ OPTIMAL"].append(card)
                    kanban_data[column_name] = [c for c in kanban_data[column_name] if c["id"] != card["id"]]
                    save_kanban_data(kanban_data)
                    st.rerun()
        
        with col4:
            # More actions dropdown
            with st.popover("⋯", help="Lainnya"):
                st.caption(f"ID: {card['id']}")
                
                # Edit form
                with st.form(key=f"edit_form_{card['id']}_{index}"):
                    new_text = st.text_area("Judul", value=card["text"], key=f"text_{card['id']}_{index}")
                    new_label = st.selectbox("Label", list(LABEL_COLORS.keys()), 
                                           index=list(LABEL_COLORS.keys()).index(card["label"]) if card["label"] in LABEL_COLORS else 0,
                                           key=f"label_{card['id']}_{index}")
                    new_priority = st.selectbox("Prioritas", list(PRIORITY_COLORS.keys()),
                                              index=list(PRIORITY_COLORS.keys()).index(card["priority"]) if card["priority"] in PRIORITY_COLORS else 1,
                                              key=f"priority_{card['id']}_{index}")
                    new_assignee = st.text_input("Assignee", value=card.get("assignee", ""), key=f"assignee_{card['id']}_{index}")
                    
                    # Handle due date
                    current_due = None
                    if card["due_date"]:
                        try:
                            current_due = datetime.strptime(card["due_date"], "%Y-%m-%d").date()
                        except:
                            current_due = datetime.now().date() + timedelta(days=7)
                    else:
                        current_due = datetime.now().date() + timedelta(days=7)
                    
                    new_due_date = st.date_input("Due Date", value=current_due, key=f"due_{card['id']}_{index}")
                    
                    col_a, col_b = st.columns(2)
                    with col_a:
                        if st.form_submit_button("💾 Simpan"):
                            kanban_data = get_kanban_data()
                            for col in kanban_data.values():
                                for c in col:
                                    if c["id"] == card["id"]:
                                        c["text"] = new_text
                                        c["label"] = new_label
                                        c["priority"] = new_priority
                                        c["assignee"] = new_assignee
                                        c["due_date"] = new_due_date.strftime("%Y-%m-%d")
                                        break
                            save_kanban_data(kanban_data)
                            st.success("Disimpan!")
                            st.rerun()
                    
                    with col_b:
                        if st.form_submit_button("🗑️ Hapus", type="secondary"):
                            kanban_data = get_kanban_data()
                            kanban_data[column_name] = [c for c in kanban_data[column_name] if c["id"] != card["id"]]
                            save_kanban_data(kanban_data)
                            st.success("Dihapus!")
                            st.rerun()
                
                # Move to column
                st.subheader("Pindah ke:")
                target_cols = list(get_kanban_data().keys())
                current_idx = target_cols.index(column_name)
                
                cols = st.columns(len(target_cols))
                for idx, (col_widget, col_name) in enumerate(zip(cols, target_cols)):
                    with col_widget:
                        if idx != current_idx:
                            if st.button(col_name[:2], key=f"move_{card['id']}_{idx}_{index}", help=f"Pindah ke {col_name}"):
                                kanban_data = get_kanban_data()
                                kanban_data[col_name].append(card)
                                kanban_data[column_name] = [c for c in kanban_data[column_name] if c["id"] != card["id"]]
                                save_kanban_data(kanban_data)
                                st.rerun()
        
        st.divider()

def render_column(column_name, kanban_data):
    """Render a single column with progress bar"""
    cards = kanban_data[column_name]
    bg_color = COLUMN_COLORS.get(column_name, "#ffffff")
    border_color = COLUMN_BORDER_COLORS.get(column_name, "#d9d9d9")
    
    # Calculate column stats
    total_cards = len(cards)
    high_priority = sum(1 for c in cards if c.get("priority") == "High")
    
    # Column header with stats
    st.markdown(f"""
    <div style="background: {bg_color}; border: 2px solid {border_color}; 
                border-radius: 12px; padding: 15px; margin-bottom: 15px; min-height: 500px;">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;">
            <h4 style="margin: 0; color: #333;">{column_name}</h4>
            <span style="background: #666; color: white; border-radius: 12px; padding: 2px 10px; font-size: 12px;">
                {total_cards}
            </span>
        </div>
    """, unsafe_allow_html=True)
    
    # Progress bar for high priority items
    if total_cards > 0:
        progress = high_priority / total_cards
        st.progress(progress)
        st.caption(f"🚨 {high_priority} prioritas tinggi")
    
    # Empty state
    if not cards:
        st.info("📭 Tidak ada kartu", icon="ℹ️")
        st.markdown("</div>", unsafe_allow_html=True)
        return
    
    # Render cards with unique keys
    for i, card in enumerate(cards):
        render_card(card, column_name, i)
    
    st.markdown("</div>", unsafe_allow_html=True)

# ============================================================
# ANALYTICS FUNCTIONS
# ============================================================
def render_analytics():
    """Render analytics dashboard"""
    st.subheader("📊 Analytics Dashboard")
    
    stats = get_card_statistics()
    kanban_data = get_kanban_data()
    
    # Key Metrics
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Kartu", stats["total_cards"])
    with col2:
        st.metric("Prioritas Tinggi", stats["high_priority"])
    with col3:
        st.metric("Terlambat", stats["overdue"])
    with col4:
        st.metric("Completion Rate", f"{stats['completion_rate']:.1f}%")
    
    # Charts row 1
    col1, col2 = st.columns(2)
    
    with col1:
        # Priority distribution
        priority_data = {
            "Prioritas": ["High", "Medium", "Low"],
            "Jumlah": [stats["high_priority"], stats["medium_priority"], stats["low_priority"]]
        }
        fig = px.pie(priority_data, values='Jumlah', names='Prioritas', 
                     title='Distribusi Prioritas',
                     color='Prioritas',
                     color_discrete_map={'High': '#ff4d4f', 'Medium': '#faad14', 'Low': '#52c41a'})
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        # Assignee distribution
        if stats["by_assignee"]:
            assignee_data = pd.DataFrame({
                "Assignee": list(stats["by_assignee"].keys()),
                "Jumlah": list(stats["by_assignee"].values())
            }).sort_values("Jumlah", ascending=False)
            
            fig = px.bar(assignee_data, x='Assignee', y='Jumlah', 
                        title='Distribusi per Assignee',
                        color='Jumlah',
                        color_continuous_scale='Blues')
            st.plotly_chart(fig, use_container_width=True)
    
    # Charts row 2
    col1, col2 = st.columns(2)
    
    with col1:
        # Burndown chart
        burndown_data = get_burndown_data()
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=burndown_data['Tanggal'], y=burndown_data['Masalah Aktif'],
                                mode='lines+markers', name='Masalah Aktif', line=dict(color='red')))
        fig.add_trace(go.Scatter(x=burndown_data['Tanggal'], y=burndown_data['Terselesaikan'],
                                mode='lines+markers', name='Terselesaikan', line=dict(color='green')))
        fig.update_layout(title='Burndown Chart (7 Hari)',
                         xaxis_title='Tanggal',
                         yaxis_title='Jumlah')
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        # Column distribution
        column_data = pd.DataFrame({
            "Kolom": list(kanban_data.keys()),
            "Jumlah": [len(cards) for cards in kanban_data.values()]
        })
        
        fig = px.bar(column_data, x='Kolom', y='Jumlah', 
                    title='Distribusi per Kolom',
                    color='Jumlah',
                    color_continuous_scale='Viridis')
        st.plotly_chart(fig, use_container_width=True)
    
    # Detailed statistics
    with st.expander("📈 Detail Statistik"):
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("**Per Label:**")
            for label, count in stats["by_label"].items():
                st.write(f"• {label}: {count}")
        
        with col2:
            st.write("**Status Due Date:**")
            st.write(f"• Terlambat: {stats['overdue']}")
            st.write(f"• Jatuh tempo hari ini: {stats['due_today']}")
            st.write(f"• Total dengan due date: {stats['overdue'] + stats['due_today']}")

# ============================================================
# SETTINGS FUNCTIONS
# ============================================================
def render_settings():
    """Render settings panel"""
    st.subheader("⚙️ Settings")
    
    with st.form("kanban_settings"):
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("**Warna Label:**")
            for label, color in LABEL_COLORS.items():
                new_color = st.color_picker(label, color, key=f"color_{label}")
                if new_color != color:
                    LABEL_COLORS[label] = new_color
        
        with col2:
            st.write("**Warna Prioritas:**")
            for priority, color in PRIORITY_COLORS.items():
                new_color = st.color_picker(priority, color, key=f"color_{priority}")
                if new_color != color:
                    PRIORITY_COLORS[priority] = new_color
        
        if st.form_submit_button("💾 Simpan Pengaturan"):
            st.success("Pengaturan disimpan!")

# ============================================================
# MAIN RENDER FUNCTION - PURE STREAMLIT VERSION
# ============================================================
def render_drag_kanban():
    # Custom CSS
    st.markdown("""
    <style>
    .kanban-card {
        background: white;
        border-radius: 8px;
        padding: 12px;
        margin-bottom: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        border-left: 5px solid;
        transition: transform 0.2s;
    }
    .kanban-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.12);
    }
    .card-header {
        margin-bottom: 10px;
    }
    .card-text {
        font-weight: 500;
        font-size: 14px;
        line-height: 1.4;
        color: #333;
    }
    .card-meta {
        display: flex;
        gap: 8px;
        margin-top: 6px;
        align-items: center;
    }
    .card-label {
        padding: 2px 8px;
        border-radius: 12px;
        font-size: 11px;
        font-weight: 500;
        color: white;
    }
    .card-priority {
        font-size: 11px;
        font-weight: 500;
    }
    .card-footer {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-top: 8px;
        font-size: 11px;
        color: #666;
    }
    .card-assignee {
        background: #f5f5f5;
        padding: 2px 8px;
        border-radius: 10px;
    }
    .card-dates {
        display: flex;
        gap: 8px;
    }
    .overdue {
        color: #ff4d4f !important;
        font-weight: bold;
    }
    .due-today {
        color: #faad14 !important;
        font-weight: bold;
    }
    .due-soon {
        color: #1890ff !important;
    }
    .stProgress > div > div > div > div {
        background-color: #1890ff;
    }
    @media (max-width: 768px) {
        .kanban-columns {
            grid-template-columns: 1fr !important;
        }
    }
    </style>
    """, unsafe_allow_html=True)
    
    st.title("📌 Kanban Board - Manajemen Jadwal Dokter")
    
    # Get current kanban data
    kanban_data = get_kanban_data()
    
    # Sidebar controls
    with st.sidebar:
        st.header("⚙️ Kontrol Kanban")
        
        # Last saved indicator
        if "last_saved" in st.session_state:
            st.caption(f"💾 Terakhir disimpan: {st.session_state['last_saved']}")
        
        # Generate cards from schedule
        if st.button("🔄 Generate dari Jadwal", use_container_width=True):
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
            new_text = st.text_input("Judul Kartu *")
            new_label = st.selectbox("Label", list(LABEL_COLORS.keys()))
            new_priority = st.selectbox("Prioritas", list(PRIORITY_COLORS.keys()))
            target_column = st.selectbox("Kolom Tujuan", list(kanban_data.keys()))
            new_assignee = st.text_input("Assignee", value="Admin")
            new_due_date = st.date_input("Due Date", value=datetime.now() + timedelta(days=7))
            
            if st.form_submit_button("Tambah Kartu", use_container_width=True):
                if new_text.strip():
                    new_card = {
                        "id": get_next_card_id(),
                        "text": new_text,
                        "label": new_label,
                        "priority": new_priority,
                        "created": datetime.now().strftime("%Y-%m-%d"),
                        "due_date": new_due_date.strftime("%Y-%m-%d"),
                        "assignee": new_assignee
                    }
                    
                    kanban_data[target_column].append(new_card)
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
                file_name=f"kanban_jadwal_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
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
        
        # Quick stats
        st.subheader("📈 Quick Stats")
        stats = get_card_statistics()
        st.write(f"• Total: {stats['total_cards']}")
        st.write(f"• Prioritas Tinggi: {stats['high_priority']}")
        st.write(f"• Terlambat: {stats['overdue']}")
        st.write(f"• Completion: {stats['completion_rate']:.1f}%")
        
        st.divider()
        
        # Reset buttons
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🗑️ Reset Default", use_container_width=True, type="secondary"):
                save_kanban_data(DEFAULT_KANBAN.copy())
                st.success("Reset berhasil!")
                st.rerun()
        
        with col2:
            if st.button("🧹 Kosongkan", use_container_width=True, type="secondary"):
                for column in kanban_data:
                    kanban_data[column] = []
                save_kanban_data(kanban_data)
                st.success("Semua kartu dihapus!")
                st.rerun()
    
    # Main content with tabs
    tab1, tab2, tab3 = st.tabs(["📋 Kanban Board", "📊 Analytics", "⚙️ Settings"])
    
    with tab1:
        # Card Mover Section
        st.subheader("🔄 Pindahkan Kartu")
        
        # Get all cards for moving
        all_cards_for_moving = []
        card_dict = {}
        for col_name, cards in kanban_data.items():
            for card in cards:
                all_cards_for_moving.append({
                    "id": card["id"],
                    "text": card["text"],
                    "current_column": col_name,
                    "full_text": f"{card['text']} ({col_name})"
                })
                card_dict[card["id"]] = card
        
        if all_cards_for_moving:
            # Select card to move
            selected_option = st.selectbox(
                "Pilih kartu untuk dipindahkan:",
                options=[card["full_text"] for card in all_cards_for_moving],
                key="card_selector"
            )
            
            if selected_option:
                # Find selected card
                selected_card = None
                for card in all_cards_for_moving:
                    if card["full_text"] == selected_option:
                        selected_card = card
                        break
                
                if selected_card:
                    col1, col2, col3 = st.columns([2, 2, 1])
                    
                    with col1:
                        st.info(f"Kartu saat ini di: **{selected_card['current_column']}**")
                    
                    with col2:
                        target_column = st.selectbox(
                            "Pindah ke kolom:",
                            options=list(kanban_data.keys()),
                            index=list(kanban_data.keys()).index(selected_card["current_column"]),
                            key="target_column"
                        )
                    
                    with col3:
                        st.write("")
                        st.write("")
                        if st.button("🚀 Pindahkan", use_container_width=True, key="move_button"):
                            if target_column != selected_card["current_column"]:
                                # Remove from current column
                                kanban_data[selected_card["current_column"]] = [
                                    c for c in kanban_data[selected_card["current_column"]] 
                                    if c["id"] != selected_card["id"]
                                ]
                                
                                # Add to target column
                                if selected_card["id"] in card_dict:
                                    kanban_data[target_column].append(card_dict[selected_card["id"]])
                                    save_kanban_data(kanban_data)
                                    st.success(f"Kartu dipindahkan ke {target_column}!")
                                    st.rerun()
                            else:
                                st.warning("Pilih kolom yang berbeda!")
        
        st.divider()
        
        # Main Kanban Board
        st.subheader("📋 Board Utama")
        
        # Responsive grid for columns
        st.markdown('<div class="kanban-columns" style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px;">', 
                   unsafe_allow_html=True)
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            render_column("⚠️ MASALAH JADWAL", kanban_data)
        
        with col2:
            render_column("🔧 PERLU PENYESUAIAN", kanban_data)
        
        with col3:
            render_column("⏳ DALAM PROSES", kanban_data)
        
        with col4:
            render_column("✅ OPTIMAL", kanban_data)
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Bulk Operations
        st.divider()
        st.subheader("⚡ Operasi Massal")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if st.button("▶️ Mulai Semua di 'MASALAH'", use_container_width=True):
                # Move all cards from MASALAH JADWAL to DALAM PROSES
                cards_to_move = kanban_data["⚠️ MASALAH JADWAL"].copy()
                kanban_data["⏳ DALAM PROSES"].extend(cards_to_move)
                kanban_data["⚠️ MASALAH JADWAL"] = []
                save_kanban_data(kanban_data)
                st.success(f"{len(cards_to_move)} kartu dipindahkan!")
                st.rerun()
        
        with col2:
            if st.button("✅ Selesaikan Semua di 'PROSES'", use_container_width=True):
                # Move all cards from DALAM PROSES to OPTIMAL
                cards_to_move = kanban_data["⏳ DALAM PROSES"].copy()
                kanban_data["✅ OPTIMAL"].extend(cards_to_move)
                kanban_data["⏳ DALAM PROSES"] = []
                save_kanban_data(kanban_data)
                st.success(f"{len(cards_to_move)} kartu diselesaikan!")
                st.rerun()
        
        with col3:
            if st.button("🗑️ Hapus Semua Kartu", use_container_width=True, type="secondary"):
                if st.checkbox("Konfirmasi hapus semua kartu"):
                    for column in kanban_data:
                        kanban_data[column] = []
                    save_kanban_data(kanban_data)
                    st.success("Semua kartu dihapus!")
                    st.rerun()
    
    with tab2:
        render_analytics()
    
    with tab3:
        render_settings()

# ============================================================
# RUN APPLICATION
# ============================================================
if __name__ == "__main__":
    render_drag_kanban()
