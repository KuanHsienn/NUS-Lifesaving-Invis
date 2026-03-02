import os
import pandas as pd
from processors.certificate_generator import generate_certificates
from processors.registration import process_registrations
from processors.booklet import parse_program_booklet
from processors.results import generate_event_results

# --- CONFIGURATION ---
SCRIPT_DIR = os.path.dirname(__file__)
DATA_DIR = os.path.join(SCRIPT_DIR, "Data")
if not os.path.exists(DATA_DIR): os.makedirs(DATA_DIR)
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "Final_Reports")
if not os.path.exists(OUTPUT_DIR): os.makedirs(OUTPUT_DIR)

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

def main():
    print("\n--- STEP 1: REGISTERING ALL TEAMS ---")
    participants, registry, event_order = process_registrations(
        SCRIPT_DIR, OUTPUT_DIR, DIV_MAP, SPECIAL_SHEETS
    )

    # --- Uncomment this step after inserting the competitors to their heats ---
    print("\n--- STEP 2: PARSING PROGRAM BOOKLET ---")
    master_list = parse_program_booklet(DATA_DIR, OUTPUT_DIR)

    # --- Uncomment this step after the invis to generate final results ---
    print("\n--- STEP 3: GENERATING FINAL RESULTS ---")
    generate_event_results(DATA_DIR, OUTPUT_DIR)

    # --- Uncomment this step after the invis to generate certs ---
    print("\n--- STEP 4: GENERATING CERTS ---")
    generate_certificates(DATA_DIR, OUTPUT_DIR)

if __name__ == "__main__":
    main()