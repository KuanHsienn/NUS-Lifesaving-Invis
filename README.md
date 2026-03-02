# 🏊‍♂️ NUS Invitational Lifesaving Automation

A modular Python automation suite designed to process championship registrations, parse programme booklets, generate dynamic event results, and export personalized certificates.

---

## 🚀 Features

- **Dynamic Column Discovery** 🔎

  - Scans Excel headers (e.g., `"Competitor No."`, `"Final Timing"`) to find columns by name rather than fixed letters, making the parsers robust to layout changes.

- **Intelligent Result Ranking** 🧠

  - **Swimming events**: averages multiple timings and ranks competitors by speed automatically.
  - **SERC events**: detects `"Simulated Emergency Response Competition"` and switches to a points-based ranking system.
  - **Validation rules**: flags rows as "Verified" only when both timing and position are present.

- **Professional Excel Formatting** ✨

  - Merges cells for multi-line competitor names for readability.
  - Dynamically adjusts row heights based on number of participants (newline counts in competitor IDs).
  - Displays missing points as `-` for a clean output.

- **Certificate Automation** 🎓
  - Replaces `{{name}}` placeholders in PPTX templates and exports to PDF via PowerPoint COM (win32com).
  - Uses explicit garbage collection, COM object deletion, and short delays to safely remove temporary PPTX files on Windows.

---

## 📁 Project Structure

```text
INVIS CODE/
├── data/                       # Input files (Excel datasets & PPTX templates)
├── Final_Reports/              # Generated outputs (Master Lists, Results, PDFs)
├── processors/
│   ├── registration.py         # Team registration and participant mapping
│   ├── booklet.py              # Parsing heats/lanes with dynamic column logic
│   ├── results.py              # Result ranking, points calculation & formatting
│   └── certificates.py         # PPTX generation and PDF export
├── Team_Line_Ups/              # Holds team registration forms
├── utils/
│   └── helpers.py              # regex, time formatting and shared utilities
├── main.py
└── README.md
```

---

## 🛠️ Setup & Usage

### Prerequisites

- **Python 3.12+**
- **Microsoft PowerPoint** (required only for PPTX->PDF export using COM on Windows)

### Required packages

```bash
pip install -r requirements.txt
```


### Running the automation

1. Drop source files into the `data/` folder (e.g., `2025 Programme Booklet.xlsx`, `volunteers.xlsx`, `certificate_template.pptx`).
2. Run the main controller:

```bash
python main.py
```

Outputs (master lists, event results, PDFs) will be saved into `Final_Reports/`.

---

## ⚙️ Configuration

- The main controller exposes two simple mapping utilities in `main.py`:
  - `DIV_MAP` — handle division naming conventions
  - `SPECIAL_SHEETS` — map sheet names for special events

These allow you to adapt the system to new input conventions without touching parser logic.

> ⚠️ Note: The PPTX->PDF conversion uses `win32com` and therefore only runs on Windows with PowerPoint installed. If running on Linux/macOS, the certificate-generation step will still produce PPTX files but cannot convert them to PDF automatically.

---

## 🧩 Implementation Notes

- Parsers do not rely on fixed Excel columns; they locate headers at runtime for resilience.
- Result calculation is modular: adding a new event type is a matter of defining a ranking strategy and registering it with the results processor.
- The certificate pipeline: generate PPTX -> use PowerPoint COM to export PDF -> delete temporary PPTX after ensuring handles have been released (`gc.collect()` + `time.sleep()` + `del` on COM objects).

---
