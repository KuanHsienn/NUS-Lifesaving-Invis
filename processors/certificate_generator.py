import pandas as pd
from pptx import Presentation
import os
import glob
import win32com.client
import time
import gc

def replace_placeholders(prs, name):
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if "{{name}}" in run.text:
                            run.text = run.text.replace("{{name}}", name)

def generate_certificates(script_dir, output_dir):
    cert_output_dir = os.path.join(output_dir, "Certificates")
    if not os.path.exists(cert_output_dir):
        os.makedirs(cert_output_dir)
        
    template_path = os.path.join(script_dir, "certificate_template.pptx")
    if not os.path.exists(template_path):
        print(f"Template not found at {template_path}. Skipping Certs.")
        return

    matches = [f for f in glob.glob(os.path.join(script_dir, "*.xlsx")) 
               if "volunteers" in os.path.basename(f).lower()]
    
    if not matches:
        print("No 'volunteers.xlsx' found for certificate generation.")
        return

    df = pd.read_excel(matches[0], header=None, names=["Name"])
    df["Name"] = df["Name"].str.strip()
    df = df.dropna(subset=["Name"]).drop_duplicates(subset="Name")

    print("Opening PowerPoint for PDF export...")
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")

    try:
        for index, row in df.iterrows():
            if index == 0: continue
            name = str(row["Name"])
            individual_names = [n.strip() for n in name.replace('\n', ',').split(',') if n.strip()]

            for single_name in individual_names:
                safe_name = single_name.replace(" ", "_").upper()
                pptx_path = os.path.join(cert_output_dir, f"{safe_name}_CERT.pptx")
                pdf_path = os.path.join(cert_output_dir, f"{safe_name}_CERT.pdf")

                if os.path.exists(pdf_path):
                    continue

                prs = Presentation(template_path)
                replace_placeholders(prs, single_name)
                prs.save(pptx_path)

                abs_pptx = os.path.abspath(pptx_path)
                abs_pdf = os.path.abspath(pdf_path)
                
                pres = powerpoint.Presentations.Open(abs_pptx, WithWindow=False)
                pres.SaveAs(abs_pdf, 32)  
                pres.Close()
                
                del pres
                gc.collect()

                time.sleep(0.5)

                os.remove(pptx_path)
                print(f"Generated PDF for: {single_name}")

    finally:
        powerpoint.Quit()
        print("Certificate generation complete.")