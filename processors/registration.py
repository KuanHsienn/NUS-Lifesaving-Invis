import pandas as pd
import re
import os
import glob
from collections import defaultdict, Counter
from utils.helpers import clean_event_code, is_team_event, get_gender_code

def process_registrations(script_dir, output_dir, div_code_map, special_team_sheets):
    registration_folder = os.path.join(script_dir, "Team_Line_Ups")
    input_files = glob.glob(os.path.join(registration_folder, "*TeamLineUp*.xlsx"))
    
    event_participants = defaultdict(list) # eg. "50m Carry": [ {...}, {...} ]
    ordered_event_names = []
    event_map_global = {}
    team_member_registry = {}
    primary_key_counter = Counter()

    # --- Data Ingestion Phase ---
    for excel_path in input_files:
        print(f"Processing Registration: {os.path.basename(excel_path)}")
        xls = pd.ExcelFile(excel_path)

        # Extract Team Code — scan column C for first cell containing a (XXX) 3-letter code
        event_list_df = xls.parse("Event list", header=None)
        team_code = "UNK"
        for val in event_list_df.iloc[:, 2].astype(str):
            m = re.search(r"\((\w{3})\)", val)
            if m:
                team_code = m.group(1)
                break

        # Map Event Codes and maintain Order
        if not event_map_global:
            event_df = event_list_df.iloc[1:, [0, 1]]
            for _, row in event_df.iterrows():
                code, name = str(row.iloc[0]).strip(), str(row.iloc[1]).strip()
                if re.match(r"^E\d+$", code):
                    event_map_global[code] = name
                    if name not in ordered_event_names:
                        ordered_event_names.append(name)

            # Add special sheets to global order
            for spec_name in set(special_team_sheets.values()):
                if spec_name not in ordered_event_names:
                    ordered_event_names.append(spec_name)

        # Process Sheets
        for sheet in [s for s in xls.sheet_names if s not in ["Event list", "Example LineUp"]]:
            df = xls.parse(sheet, header=None)
            gender_code = "X" if sheet in special_team_sheets else get_gender_code(sheet)
            division_cell = str(df.iloc[2, 0]).strip().upper()
            div_code = next((v for k, v in div_code_map.items() if k in division_cell), "?")

            if sheet in special_team_sheets:
                process_special_sheet(df, sheet, team_code, div_code, special_team_sheets, team_member_registry, primary_key_counter, event_participants)
            else:
                process_regular_sheet(df, sheet, team_code, div_code, gender_code, event_map_global, team_member_registry, primary_key_counter, event_participants)

    output_path = os.path.join(output_dir, "Parsed_Event_List.xlsx")
    _save_to_excel(output_path, team_member_registry, event_participants, ordered_event_names)

    print(f"Registration processing complete. File saved to: {output_path}")
    return event_participants, team_member_registry, ordered_event_names, event_map_global

def _save_to_excel(output_path, registry, participants, ordered_names):
    div_priority = {'Y': 0, 'B': 1, 'A': 2, 'O': 3}

    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        wrap_format = writer.book.add_format({'text_wrap': True})

        # 1. Participants Sheet
        df_members = pd.DataFrame([{"Name": n, "ID": p} for n, p in registry.items()]).sort_values("ID")
        df_members.to_excel(writer, sheet_name="All Participants", index=False)
        apply_formatting(writer.sheets["All Participants"], df_members, wrap_format)

        # 2. Event Sheets
        for event_name in ordered_names:
            if event_name not in participants: continue
            rows = participants[event_name]

            if is_team_event(event_name):
                grouped = defaultdict(list)
                for r in rows:
                    key = (r["Inst"], r["Div"], r["Team"])
                    grouped[key].append((r["No."], r["Competitor"]))
                
                output_rows = []
                for (inst, div, team), members in grouped.items():
                    ids, names = zip(*members)
                    output_rows.append({
                        "COMPETITORS": " / ".join(names),
                        "No.": "\n".join(ids),
                        "INST": inst,
                        "DIV": div,
                        "TEAM CODE": team,
                        "_priority": div_priority.get(div, 99)
                    })
                df_event = pd.DataFrame(output_rows)
            else:
                df_event = pd.DataFrame(rows)
                df_event["_priority"] = df_event["Div"].map(lambda x: div_priority.get(x, 99))

            # Sorting logic for easier copy
            df_event = df_event.sort_values(by=["_priority", "INST" if is_team_event(event_name) else "Inst"])
            df_event = df_event.drop(columns=["_priority"])

            sheet_name = event_name[:31]
            df_event.to_excel(writer, sheet_name=sheet_name, index=False)
            apply_formatting(writer.sheets[sheet_name], df_event, wrap_format)

def apply_formatting(worksheet, df, wrap_format):
    for i, col in enumerate(df.columns):
        max_len = max(df[col].astype(str).map(len).max(), len(str(col)))
        worksheet.set_column(i, i, max_len + 2)
    for row_num in range(len(df) + 1):
        worksheet.set_row(row_num, None, wrap_format)

def process_special_sheet(df, sheet, team_code, div_code, special_team_sheets, registry, counter, participants):
    event_name = special_team_sheets[sheet]
    current_team = []
    team_counter = 1
    
    for row_idx in range(4, len(df)):
        name = df.iloc[row_idx, 2]
        if pd.isna(name) or str(name).strip().lower() == "name": continue
        name = str(name).strip().upper()
        
        pk = get_or_create_pk(name, (team_code, "X", div_code), registry, counter, div_code)
        if not pk: continue

        current_team.append({"Competitor": name, "No.": pk, "Inst": team_code, "Div": div_code, "Team": f"{sheet}_T{team_counter}"})
        if len(current_team) == 4:
            participants[event_name].extend(current_team)
            current_team = []
            team_counter += 1
    if current_team: 
        participants[event_name].extend(current_team)

def process_regular_sheet(df, sheet, team_code, div_code, gender_code, event_map, registry, counter, participants):
    events = df.iloc[3, 1:]
    open_group_counters = defaultdict(int)
    open_current_group = {}

    for row_idx in range(3, len(df)):
        first_four = [str(df.iloc[row_idx, i]).strip().lower() for i in range(1, 5)]
        if first_four == ["no.", "name", "date of birth", "bm ref no."]:
            for col_idx in range(1, len(events)+1):
                raw_event_code = str(df.iloc[3, col_idx]).replace(" ", "")
                e_name = event_map.get(clean_event_code(raw_event_code))
                if e_name and is_team_event(e_name) and div_code == "O":
                    base_key = (e_name, team_code, div_code, raw_event_code)
                    open_group_counters[base_key] += 1
                    open_current_group[base_key] = f"{raw_event_code}_G{open_group_counters[base_key]}"
            continue

        name = df.iloc[row_idx, 2]
        if pd.isna(name): continue
        name = str(name).strip().upper()

        pk = get_or_create_pk(name, (team_code, gender_code, div_code), registry, counter, div_code)
        if not pk: continue

        for col_idx in range(5, len(events)+1):
            if str(df.iloc[row_idx, col_idx]).strip().lower() == "x":
                raw_event_code = str(df.iloc[3, col_idx]).replace(" ", "")
                e_name = event_map.get(clean_event_code(raw_event_code))
                if not e_name: continue

                group_id = raw_event_code
                if is_team_event(e_name) and div_code == "O":
                    group_id = open_current_group.get((e_name, team_code, div_code, raw_event_code), f"{raw_event_code}_FB")

                participants[e_name].append({"Competitor": name, "No.": pk, "Inst": team_code, "Div": div_code, "Team": group_id})

def get_or_create_pk(name, group_key, registry, counter, div_code):
    if name in registry: return registry[name]
    if div_code == "O" or counter[group_key] < 6:
        counter[group_key] += 1
        pk = f"{group_key[0]}{group_key[1]}{group_key[2]}{counter[group_key]:02d}"
        registry[name] = pk
        return pk
    print(f"Warning: participant limit reached for {group_key} — '{name}' not registered")
    return None