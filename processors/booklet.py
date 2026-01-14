import pandas as pd
import re
import os
import glob

def parse_program_booklet(script_dir, output_dir):
    program_files = [f for f in glob.glob(os.path.join(script_dir, "*.xlsx")) 
                     if "programme" in os.path.basename(f).lower()]
    if not program_files: 
        print("⚠️ No program booklet found.")
        return []

    booklet = pd.ExcelFile(program_files[0])
    master_rows = []
    serial_number = 1
    heat_pattern = re.compile(r"finals\s+(\d+)\s+of\s+(\d+)", re.IGNORECASE)

    # --- Process Sheets ---
    for sheet_name in booklet.sheet_names:
        if not re.fullmatch(r"E([0-9]*)", sheet_name, re.IGNORECASE): 
            continue
        
        df = booklet.parse(sheet_name, header=None)
        event_code = sheet_name.upper()
        
        # Determine Title Cell logic
        event_name = "Unknown Event" 

        # Scan the first 3 rows to find the title
        for r_idx in range(3):
            first_cell = str(df.iloc[r_idx, 0]).strip()
            
            # Titles look like "Event 23: Mixed Pool Lifesaver Relay" 
            if first_cell and first_cell.lower() != 'nan' and len(first_cell) > 5:
                if ":" in first_cell:
                    event_name = first_cell.split(":", 1)[-1].strip()
                break
        
        current_heat, capture_rows = "H1", False
        cols = {}

        for idx, row in df.iterrows():
            row_str = [str(c).strip().lower() for c in row]
            # Heat Detection: Reset capture_rows to False to wait for the next header
            for cell in row:
                if isinstance(cell, str) and heat_pattern.search(cell):
                    current_heat = f"H{heat_pattern.search(cell).group(1)}"
                    capture_rows = False

            # Header Detection (Lane, Competitors, No., Inst)
            if sum(1 for kw in ["lane", "competitors", "no.", "inst"] if kw in row_str) >= 3:
                cols = {kw: row_str.index(kw) for kw in ["lane", "competitors", "no.", "inst", "team"] if kw in row_str}
                capture_rows = True
                continue

            # Row Data Capture
            if capture_rows and not all(pd.isna(c) for c in row):
                name = str(row[cols['competitors']]).strip()
                if not name or name.lower() == "competitors" or not name.isupper(): 
                    continue
                row_is_serc = "SIMULATED EMERGENCY RESPONSE COMPETITION" in event_name.upper()

                master_rows.append({
                    "S/N": serial_number,
                    "Event Name": event_name,
                    "Event Code": event_code,
                    "Heat No.": current_heat,
                    "Lane No.": str(row[cols['lane']]) if 'lane' in cols else "",
                    "Competitor Name": name,
                    "Team": str(row[cols['inst']]) if 'inst' in cols else "",
                    "Div": "SERC" if row_is_serc else (str(row[cols['team']]) if 'team' in cols else ""),
                    "Competitor No.": str(row[cols['no.']]) if 'no.' in cols else "",
                    "Timing 1": "",
                    "Timing 2": "",
                    "Timing 3": "",
                    "Final Timing": None,
                    "Final Backend Timing": None, 
                    "Position": "",
                    "Points": None,
                    "Verified": None
                })
                serial_number += 1

    # --- Save to Excel with Formatting ---
    if master_rows:
        output_path = os.path.join(output_dir, "Program_Master_List.xlsx")
        df_master = pd.DataFrame(master_rows)
        
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df_master.to_excel(writer, sheet_name="Master Event List", index=False)
            workbook = writer.book
            worksheet = writer.sheets["Master Event List"]
            
            # --- DYNAMIC COLUMN FINDER ---
            def get_col_idx(col_name):
                return df_master.columns.get_loc(col_name)

            def get_col_letter(col_name):
                return chr(65 + df_master.columns.get_loc(col_name))

            idx_t1 = get_col_idx("Timing 1")
            idx_final = get_col_idx("Final Timing")
            idx_backend = get_col_idx("Final Backend Timing")
            idx_pos = get_col_idx("Position")
            idx_pts = get_col_idx("Points")
            idx_ver = get_col_idx("Verified")

            L_T1 = get_col_letter("Timing 1")
            L_T3 = get_col_letter("Timing 3")
            L_DIV = get_col_letter("Div")
            L_FINAL = get_col_letter("Final Timing")
            L_BACK = get_col_letter("Final Backend Timing")
            L_POS = get_col_letter("Position")
            L_PTS = get_col_letter("Points")
            L_CODE = get_col_letter("Event Code")
            L_EVENT = get_col_letter("Event Name")

            # Formats
            num_format = workbook.add_format({
                'font_name': 'Calibri', 
                'font_size': 11, 
                'num_format': '0', 
                'valign': 'vcenter'
            })
            time_format = workbook.add_format({
                'font_name': 'Calibri', 
                'font_size': 11, 
                'num_format': 'm"min" s"s"  .000"ms"',
                'valign': 'vcenter'
            })
            
            # Column Letter Mapping for Formulas
            # J(9)=Timing1, K=Timing2, L=Timing3, M=FinalTiming, N(13)=FinalBackend, O=Position, P=Points, Q(16)=Verified

            worksheet.set_column(idx_t1, idx_backend, None, time_format) 

            for i in range(len(df_master)):
                current_event_name = str(df_master.iloc[i]['Event Name']).upper()
                is_serc = "SIMULATED EMERGENCY RESPONSE COMPETITION" == current_event_name
                excel_row = i + 2 # 1-based index + header
                
                # Verified -> Check if N and O are not blank
                worksheet.write_formula(
                    i + 1, 
                    idx_ver, 
                    f'=IF(AND({L_BACK}{excel_row}<>"", {L_POS}{excel_row}<>""), 1, "")', 
                    num_format
                )

                if is_serc:
                    # Serc Points Ranking
                    serc_rank_formula = (
                        f'=IF(NOT(ISNUMBER({L_PTS}{excel_row})), "", '
                        f'COUNTIFS('
                        f'{L_PTS}$2:{L_PTS}$999, ">"&{L_PTS}{excel_row}, '
                        f'{L_EVENT}$2:{L_EVENT}$999, "SIMULATED EMERGENCY RESPONSE COMPETITION"'
                        f')+1)'
                    )

                    worksheet.write_formula(i+1, idx_backend, serc_rank_formula, num_format)
                    worksheet.write_formula(i+1, idx_pos, serc_rank_formula, num_format)
                else:
                    # Final Timing (Average of Timing 1, 2, 3) -> J, K, L
                    worksheet.write_formula(i+1, idx_final, f"=AVERAGE({L_T1}{excel_row}:{L_T3}{excel_row})", time_format)

                    # Final Backend Timing -> matches Final Timing
                    worksheet.write_formula(i+1, idx_backend, f"={L_FINAL}{excel_row}", time_format)

                    points_formula = (
                        f'=IF(OR({L_DIV}{excel_row}="O", NOT(ISNUMBER({L_FINAL}{excel_row}))), "", '
                        f'IF('
                        f'COUNTIFS({L_CODE}$2:{L_CODE}$999, {L_CODE}{excel_row}, {L_DIV}$2:{L_DIV}$999, {L_DIV}{excel_row}, {L_FINAL}$2:{L_FINAL}$999, "<"&{L_FINAL}{excel_row})+1=1, 9, '
                        f'IF('
                        f'COUNTIFS({L_CODE}$2:{L_CODE}$999, {L_CODE}{excel_row}, {L_DIV}$2:{L_DIV}$999, {L_DIV}{excel_row}, {L_FINAL}$2:{L_FINAL}$999, "<"&{L_FINAL}{excel_row})+1<=8, '
                        f'10-(COUNTIFS({L_CODE}$2:{L_CODE}$999, {L_CODE}{excel_row}, {L_DIV}$2:{L_DIV}$999, {L_DIV}{excel_row}, {L_FINAL}$2:{L_FINAL}$999, "<"&{L_FINAL}{excel_row})+1)-1, '
                        f'""'
                        f')))'
                    )

                    worksheet.write_formula(i+1, idx_pts, points_formula, num_format)

    print(f"✅ Master List with Live Formulas created: {output_path}")
    return master_rows