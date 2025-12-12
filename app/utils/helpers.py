# app/utils/helpers.py
def chunk_list(lst, n):
    """Bagi list menjadi n-chunks (helper kecil)."""
    for i in range(0, len(lst), n):
        yield lst[i:i+n]
