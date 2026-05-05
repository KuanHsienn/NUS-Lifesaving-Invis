# NUS Invitational Lifesaving Championship — Pipeline

A Python automation suite that handles team registration, heat seeding, programme booklet generation, live results parsing, and certificate export for the NUS Invitational Lifesaving Championship.

---

## Pipeline Overview

The pipeline runs in five sequential steps controlled by `main.py`:

| Step  | Purpose                                             | When to run              |
| ----- | --------------------------------------------------- | ------------------------ |
| **0** | Update master seed from previous year's results     | After each championship  |
| **1** | Parse team line-up files, register all participants | Before each championship |
| **2** | Generate seeded programme booklet                   | Before each championship |
| **3** | Parse filled-in programme booklet → master list     | After the event          |
| **4** | Generate final ranked results                       | After the event          |

---

## Project Structure

```
INVIS CODE/
├── Data/                           # Input files
│   ├── Master_Seed_Timing.xlsx     # Cumulative historical seed times (auto-managed)
│   ├── NUS Invitational ... Results.xlsx   # Previous/current year results
│   └── <Year> Programme Booklet.xlsx       # Filled-in booklet (post-event)
├── Final_Reports/                  # All generated outputs
│   ├── Parsed_Event_List.xlsx      # Registered participants per event
│   ├── Seeded_Programme_Booklet.xlsx
│   ├── Program_Master_List.xlsx    # Timing entry sheet with live formulas
│   └── Event_Results_Final.xlsx
├── Team_Line_Ups/                  # TeamLineUp Excel files from each institution
├── processors/
│   ├── registration.py             # Team registration and participant mapping
│   ├── heat_seeding.py             # Master seed I/O and seeded booklet generation
│   ├── booklet.py                  # Post-event booklet parsing
│   └── results.py                  # Final results ranking and formatting
├── utils/
│   └── helpers.py                  # Shared utilities (event type detection, gender codes)
├── main.py                         # Pipeline controller
└── requirements.txt
```

---

## Setup

**Requirements:** Python 3.12+

```bash
pip install -r requirements.txt
```

---

## Configuration (`main.py`)

```python
CHAMPIONSHIP_YEAR = 2025        # Year of the current championship
                                # Step 0 uses results from < this year for seeding
                                # Step 4 looks for results from this year

DIV_MAP = {                     # Maps division labels in the TeamLineUp files
    "DIVISION A": "A",          # to single-letter codes used throughout the pipeline
    "DIVISION B": "B",
    "DIVISION OPEN": "O",
    "DIVISION YOUTH": "Y"
}

SPECIAL_SHEETS = {              # TeamLineUp sheet names that don't follow the
    "MIXED POOL LIFESAVER RELAY_DIVA": "Mixed Pool Lifesaver Relay",  # standard layout
    "MIXED POOL LIFESAVER RELAY_DIVB": "Mixed Pool Lifesaver Relay",
    "MIXED POOL LIFESAVER RELAY_OPEN": "Mixed Pool Lifesaver Relay",
    "SERC": "SERC"
}
```

In `heat_seeding.py`:

```python
SHOW_SEED_TIMES = True          # Shows historical seed times in the SEED column.
                                # Comment out before printing the final booklet.
```

---

## Input File Requirements

### TeamLineUp files (`Team_Line_Ups/`)

- Filename must contain `"TeamLineUp"`.
- Must have an `"Event list"` sheet with event codes (`E1`, `E2`, ...) in column A and event names in column B.
- Institution team code must appear as a 3-letter code in parentheses (e.g., `(NUS)`) somewhere in column C of the `"Event list"` sheet.
- Each event sheet row marks participation with `"X"` in the relevant column.

### Results file (`Data/`)

- Filename must contain both `"invitational"` and `"results"` (case-insensitive).
- Must include the year (e.g., `2025`) in the filename.
- Must have a `"Bout Timings"` sheet with columns: `Event Name`, `Competitor Name`, `Team`, `Div`, `Final Backend Timing`.

### Programme booklet (`Data/`)

- Filename must contain `"programme"` (case-insensitive).
- Sheet names must match the event names used in registration.

---

## Master Seed

`Data/Master_Seed_Timing.xlsx` is the cumulative record of historical best times used for lane seeding.

- **Individual events**: keyed by `(Event Name, Competitor Name)` — stores each athlete's personal best.
- **Relay/team events**: keyed by `(Event Name, '<TEAM> <DIV>')`, e.g. `NUS A` — stores the institution's fastest relay time per division.
- Only updated if the new time is strictly faster than the existing record.
- Carries forward across years — never overwritten, only appended/updated.

**Seeding logic:**

- All divisions sorted Y → B → A → O; within each division, unseeded first then slowest → fastest.
- Heats distributed equally across the full field (no per-division imbalance).
- Within each heat: fastest → centre lane (lane 4 in a 10-lane pool), spiralling outward `[4,5,3,6,2,7,1,8,0,9]`.

---

## Adding a New Event

1. Add the event code and name to the `"Event list"` sheet in the TeamLineUp file.
2. Mark participating athletes with `"X"` in the corresponding column.
3. If it is a relay/team event whose name does not contain `relay`, `line`, `emergency`, `throw`, or `serc`, add the keyword to `is_team_event()` in `utils/helpers.py`.

No other code changes are required.
