# ALTERNATIF: Ganti dengan library yang compatible
# streamlit-dnd tidak tersedia, gunakan alternatif

# Hapus import streamlit-dnd dan ganti dengan:
from streamlit_sortables import sort_items

# Atau gunakan native Streamlit dengan custom solution:
def simple_drag_drop():
    """Simple drag & drop simulation using selectboxes"""
    kanban_data = get_kanban_data()
    
    st.subheader("🔄 Pindahkan Kartu")
    
    # Get all cards with their current column
    card_options = []
    for col_name, cards in kanban_data.items():
        for card in cards:
            card_options.append({
                "id": card["id"],
                "text": card["text"],
                "current_column": col_name
            })
    
    if card_options:
        selected = st.selectbox(
            "Pilih kartu:",
            options=card_options,
            format_func=lambda x: f"{x['text']} ({x['current_column']})"
        )
        
        if selected:
            col1, col2 = st.columns(2)
            with col1:
                target_col = st.selectbox(
                    "Pindah ke:",
                    options=list(kanban_data.keys()),
                    index=list(kanban_data.keys()).index(selected["current_column"])
                )
            
            with col2:
                if st.button("🚀 Pindahkan"):
                    if target_col != selected["current_column"]:
                        # Remove from current column
                        kanban_data[selected["current_column"]] = [
                            c for c in kanban_data[selected["current_column"]] 
                            if c["id"] != selected["id"]
                        ]
                        
                        # Find card and add to target
                        all_cards = get_all_cards()
                        card_to_move = next(
                            (c for c in all_cards if c["id"] == selected["id"]),
                            None
                        )
                        
                        if card_to_move:
                            kanban_data[target_col].append(card_to_move)
                            save_kanban_data(kanban_data)
                            st.success("Berhasil dipindahkan!")
                            st.rerun()
