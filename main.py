import os
import re
import glob
import datetime
from processors.certificate_generator import generate_certificates
from processors.registration import process_registrations
from processors.booklet import parse_program_booklet
from processors.results import generate_event_results
from processors.heat_seeding import (
    update_master_seeds_from_results,
    generate_seeded_booklet,
)

# --- CONFIGURATION ---
SCRIPT_DIR = os.path.dirname(__file__)
DATA_DIR = os.path.join(SCRIPT_DIR, "Data")
if not os.path.exists(DATA_DIR): os.makedirs(DATA_DIR)
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "Final_Reports")
if not os.path.exists(OUTPUT_DIR): os.makedirs(OUTPUT_DIR)

MASTER_SEED_PATH = os.path.join(DATA_DIR, "Master_Seed_Timing.xlsx")

# Set to the year of the championship being processed.
CHAMPIONSHIP_YEAR = datetime.datetime.now().year

DIV_MAP = {
    "DIVISION A": "A",
    "DIVISION B": "B", 
    "DIVISION OPEN": "O",
    "DIVISION YOUTH": "Y"
}

SPECIAL_SHEETS = {
    "MIXED POOL LIFESAVER RELAY_DIVA": "Mixed Pool Lifesaver Relay",
    "MIXED POOL LIFESAVER RELAY_DIVB": "Mixed Pool Lifesaver Relay",
    "MIXED POOL LIFESAVER RELAY_OPEN": "Mixed Pool Lifesaver Relay",
    "SERC": "SERC"
}

def _find_results_file():
    """Return the most recent results file from before CHAMPIONSHIP_YEAR for master seed seeding."""
    matches = []
    for f in glob.glob(os.path.join(DATA_DIR, "*.xlsx")):
        name = os.path.basename(f).lower()
        if "invitational" not in name or "results" not in name:
            continue
        m = re.search(r"(20\d{2})", os.path.basename(f))
        if m and int(m.group(1)) < CHAMPIONSHIP_YEAR:
            matches.append(f)
    if not matches:
        return None
    chosen = sorted(matches)[-1]
    print(f"  Found: {os.path.basename(chosen)}")
    return chosen


def main():
    # --- STEP 0: Update master seed (post-event) ---------------------------
    print("\n--- STEP 0: UPDATE MASTER SEED TIMING (post-event) ---")
    results_file = _find_results_file()
    if results_file:
        update_master_seeds_from_results(results_file, MASTER_SEED_PATH)
    else:
        print("  No results file found - skipping master seed update.")

    # --- STEP 1: Register this year's teams --------------------------------
    print("\n--- STEP 1: REGISTERING ALL TEAMS ---")
    participants, registry, event_order, event_map = process_registrations(
        SCRIPT_DIR, OUTPUT_DIR, DIV_MAP, SPECIAL_SHEETS
    )

    # --- STEP 2: Generate seeded programme booklet -------------------------
    print("\n--- STEP 2: GENERATING SEEDED PROGRAMME BOOKLET ---")
    booklet_out = os.path.join(OUTPUT_DIR, "Seeded_Programme_Booklet.xlsx")
    generate_seeded_booklet(MASTER_SEED_PATH, participants, booklet_out, event_map)

    # --- STEP 3: Parse filled-in programme booklet -------------------------
    print("\n--- STEP 3: PARSING PROGRAM BOOKLET ---")
    parse_program_booklet(DATA_DIR, OUTPUT_DIR)

    # --- STEP 4: Generate final results ------------------------------------
    print("\n--- STEP 4: GENERATING FINAL RESULTS ---")
    generate_event_results(DATA_DIR, OUTPUT_DIR, year=CHAMPIONSHIP_YEAR)

    # # --- STEP 5: Generate certificates ------------------------------------
    # print("\n--- STEP 5: GENERATING CERTS ---")
    # generate_certificates(DATA_DIR, OUTPUT_DIR)

if __name__ == "__main__":
    main()