import re
import pandas as pd
from datetime import time


class TimeParser:
    @staticmethod
    def parse(time_str):
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
