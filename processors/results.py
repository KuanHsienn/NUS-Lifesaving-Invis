import pandas as pd
import os
import glob

def format_as_min_sec_ms(td):
    if isinstance(td, str):
        return td.strip().upper()
    if pd.isna(td) or td is pd.NaT:
        return ""
    comp = td.components
    return f"{comp.minutes}min {comp.seconds}s {comp.milliseconds}ms"

def parse_final_backend_timing(value):
    try:
        return pd.to_timedelta(value)
    except Exception:
        if isinstance(value, str) and value.strip().isalpha():
            return value.strip().upper()
        return pd.NaT

def generate_event_results(script_dir, output_dir):
    matches = [f for f in glob.glob(os.path.join(script_dir, "*.xlsx")) 
               if "invitational" in os.path.basename(f).lower()]
    
    if not matches:
        print("⚠️ No Invitational results found.")
        return

    df_raw = pd.read_excel(matches[0])
    df_raw.columns = [str(c).strip() for c in df_raw.columns]

    # --- DYNAMIC COLUMN FINDER ---
    def get_col(name): 
        return df_raw.columns.get_loc(name)

    # Map indices dynamically
    idx_name = get_col("Competitor Name")
    idx_no   = get_col("Competitor No.")
    idx_team = get_col("Team")
    idx_div  = get_col("Div")
    idx_heat = get_col("Heat No.")
    idx_lane = get_col("Lane No.")
    idx_time = get_col("Final Backend Timing")
    idx_pts  = get_col("Points")
    idx_ev_n = get_col("Event Name")

    # Clean timing data
    df_raw["Final Backend Timing"] = df_raw["Final Backend Timing"].apply(parse_final_backend_timing)
    df = df_raw[df_raw["Event Code"].notna() & df_raw["Final Backend Timing"].notna()].copy()

    output_path = os.path.join(output_dir, "Event_Results_Final.xlsx")
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        workbook = writer.book
        center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'border': 1})
        header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D3D3D3', 'border': 1})

        for event_code, group in df.groupby("Event Code", sort=False):
            # Sorting Logic
            is_valid = group["Final Backend Timing"].apply(lambda x: isinstance(x, pd.Timedelta))
            ranked = pd.concat([group[is_valid].sort_values("Final Backend Timing"), group[~is_valid]], ignore_index=True)
            
            sheet_name = str(event_code)[:31]
            event_name_text = str(ranked.iloc[0, idx_ev_n]).upper()
            is_serc = "SIMULATED EMERGENCY RESPONSE COMPETITION" in event_name_text

            worksheet = workbook.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = worksheet

            # Header Setup
            worksheet.write(4, 0, "POSITION", header_format)
            worksheet.merge_range(4, 1, 4, 7, "COMPETITORS", header_format)

            if is_serc:
                # SERC Headers
                worksheet.write(4, 8, "INST", header_format)
                worksheet.write(4, 9, "POINTS", header_format)
            else:
                # Normal Headers
                headers = ["NO.", "INST", "TEAM", "HEAT", "LANE", "TIME", "POINTS"]
                for i, text in enumerate(headers):
                    worksheet.write(4, 8 + i, text, header_format)

            # Data Writing Loop
            for i, row in ranked.iterrows():
                excel_row = 5 + i
                pos = i + 1
                
                raw_pts = row.iloc[idx_pts]
                competitor_no = row.iloc[idx_no]
                
                display_pts = "-" if pd.isna(raw_pts) else raw_pts

                if is_serc:
                    worksheet.write(excel_row, 0, pos, center_format)
                    worksheet.merge_range(excel_row, 1, excel_row, 7, row.iloc[idx_name], center_format)
                    worksheet.write(excel_row, 8, row.iloc[idx_team], center_format)
                    worksheet.write(excel_row, 9, display_pts, center_format) 
                    worksheet.set_row(excel_row, 30)
                else:
                    swim_data = [
                        row.iloc[idx_no], 
                        row.iloc[idx_team], 
                        row.iloc[idx_div],
                        row.iloc[idx_heat], 
                        row.iloc[idx_lane], 
                        format_as_min_sec_ms(row.iloc[idx_time]), 
                        display_pts 
                    ]
                    
                    sanitized_swim_data = [("" if pd.isna(val) else val) for val in swim_data]

                    worksheet.write(excel_row, 0, pos, center_format)
                    worksheet.merge_range(excel_row, 1, excel_row, 7, row.iloc[idx_name], center_format)
                    worksheet.write_row(excel_row, 8, sanitized_swim_data, center_format)

                    num_ids = str(competitor_no).count('\n') + 1
                    worksheet.set_row(excel_row, 20 + (15 * (num_ids - 1)))

            # Column Widths
            worksheet.set_column(0, 0, 10)   
            worksheet.set_column(1, 7, 10)   
            worksheet.set_column(8, 12, 10)  
            worksheet.set_column(13, 13, 18)
            worksheet.set_column(14, 14, 10) 

    print(f"✅ Results Generated using Dynamic Indexing: {output_path}")