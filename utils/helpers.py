import re
import pandas as pd

def clean_event_code(code: str) -> str:
    """Extracts E### from a string."""
    match = re.match(r"(E\d+)", str(code))
    return match.group(1) if match else str(code)

def is_team_event(event_name: str) -> bool:
    """Checks if the event name contains relay/team keywords."""
    keywords = ["relay", "line", "emergency", "throw", "serc"]
    return any(kw in str(event_name).lower() for kw in keywords)

def get_gender_code(sheet_name: str) -> str:
    """Determines gender code based on sheet name."""
    name = sheet_name.lower()
    if "women" in name:
        return "W"
    elif "men" in name:
        return "M"
    return "U" 

def format_as_min_sec_ms(td, position=None):
    """Formats pandas Timedelta into a readable swim time string."""
    # If it's already a string (like 'DQ' or 'DNS'), just return it
    if isinstance(td, str):
        return td.upper()
    if pd.isna(td) or not hasattr(td, 'components'):
        return ""
    
    c = td.components
    return f"{c.minutes}min {c.seconds}s {c.milliseconds}ms"