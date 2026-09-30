# Rail Mill Optimizer — Django

A local-first Python/Django application designed for Rail Mill raw material planning, casting optimization, BTR (Billet/Bloom to Rolling) campaign processing, rolling plan calculation, embossing sequence management, and master data configuration.

It is built around Rail Mill plant workbooks and metallurgical reference rules, intentionally utilizing pure Django, SQLite, and Vanilla JS/HTML5 without Node.js, npm, React, or heavy frontend build tools.

---

## Key Features & Modules

### 1. Demand & Order Balance Management
- Ingest SAP order balance exports and raw `input casting plan.xlsm` workbooks (`PLAN...` sheet).
- Parse SO numbers, items, planning sections, external grades, order quantities, balance quantities, delivery dates, and cutting lengths.
- Multi-select demand filtering with net demand calculation and campaign grouping.

### 2. Heuristic Yield Loss & Casting Plan Engine
- **Configurable 4-Level Yield Loss Hierarchy**:
  1. Planning Section + Section Weight + RM Family
  2. Planning Section + RM Family
  3. RM Family
  4. Global Fallback
- Configurable base strand length, head crop, tail crop, cutting loss, and process loss percentages.
- Batch rounding strategies (Raw Material Tonnage, Finished Tonnage, or Strand Count).
- Multi-sheet Excel casting plan exporter writing to `SBB`, `250sq`, and `165SQ` format templates.

### 3. BTR Campaign Generator
- Upload `.xlsx` or `.xlsm` BTR (Billet/Bloom to Rolling) report files.
- Planning Material selection with automated header normalization.
- Zero-length cutting length resolution using min/max section tolerance rules.
- **Category-Agnostic Sorting Hierarchy**:
  1. Thickness (Ascending)
  2. Cutting Length (Descending)
  3. Number of Pieces (Descending)
  4. Embossing Standard (A to Z)
- Generates macro-enabled CAMP campaign workbooks using `CAMP PLAN FORMAT.xlsm`, preserving embedded VBA projects.

### 4. Rolling Plan Generator
- Automated Rolling Section construction (combining planning material, angle thickness, and standard suffixes).
- Raw Material size resolution (`BL_165X165=214`, `BL_250X250=490`, `BL_130X130=132`, `BB_355X280X90=480` kg/m).
- External Grade to Raw Grade matching with Higher Section vs Lighter Section resolution (>300 kg/m weight threshold).
- Final Plan sorting (STD → Raw Material → Embossing → Thickness / Section Weight).
- Dynamic section subtotal insertion (Rolling Qty & Bloom 200x200 PCS) with continuous cumulative Bloom tracking (with 7% RM allowance).
- Automated Cooling Bed length calculation with iterative divisor adjustment (maintaining maximum bed length strictly below 87 m).
- Outputs macro-enabled workbooks containing `Final Plan`, `Grade Mapping`, `Raw Material Mapping`, and `Plan Assumptions`.

### 5. Embossing Plan & Roll Shop Sequence Integration
- Embossing plan generation and sequence tracking.
- Roll Shop tool sequence setup and submission workflows.
- SAP integration portal simulation and campaign assignment interface.

### 6. Technical Master Data & Reference Rules
- **Root Radius Master (M-1)**: Import and manage Root Radius rules across planning materials, thickness, section weight, and standards.
- Section Rules, Grade Maps, LFA Batch Rules, Section Weight Maps, and Billet Length formula parameters.

---

## Installation & Setup Guide

### Prerequisites
- Python 3.11 or higher installed on Windows.
- PowerShell or Command Prompt.

### Step-by-Step Setup

1. **Clone / Open the repository**:
   Open a terminal in the project directory (`Rail_Mill_optimizer`).

2. **Create and Activate Virtual Environment**:
   ```powershell
   py -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

3. **Install Dependencies**:
   ```powershell
   pip install -r requirements.txt
   ```

4. **Apply Database Migrations**:
   ```powershell
   python manage.py migrate
   ```

5. **Import Master Data & Reference Workbooks**:
   ```powershell
   # Import core Rail Mill reference workbooks, grade maps, section rules & historical workbooks
   python manage.py import_spm_data

   # Import M-1 Root Radius master data
   python casting_plan/import_m1.py
   ```

6. **Start Development Server**:
   ```powershell
   python manage.py runserver 127.0.0.1:8000
   ```
   *Alternatively, double-click `run.bat` to launch automatically.*

7. **Access the Web Application**:
   Navigate to `http://127.0.0.1:8000/` in your web browser.

---

## Application Workflows

### Workflow 1: Casting Plan Generation
1. Go to **Demand / Orders** (`/data/`).
2. Upload `input casting plan.xlsm` or current SAP export.
3. Review order items, filter by section/grade, select desired lines, and click **Generate Casting Plan**.
4. Download the structured multi-sheet Excel file (`SBB`, `250sq`, `165SQ`).

### Workflow 2: BTR Upload → Campaign → Rolling Plan
1. Open **Upload BTR** (`/upload-btr/`).
2. Upload a BTR `.xlsx` / `.xlsm` file.
3. Select the **Planning Material** to extract.
4. Click **Generate Campaign File** to create the CAMP campaign file.
5. In **Campaign History**, select a plan date and optional camp number, then click **Generate Rolling Plan**.
6. Download the generated macro-enabled rolling plan workbook.

### Workflow 3: Master Data & Technical Rules
1. Open **Master Data** (`/master-data/`).
2. View and edit Section Rules, Grade Mappings, LFA Batch Details, Yield Loss Rules, and M-1 Root Radius settings.
3. Configuration changes immediately apply to subsequent optimizer and rolling plan calculations.

---

## Included Reference Data Sources (`data/reference/`)

| File / Directory | Description |
|---|---|
| `RAW MATERIAL LENGTH TOLERANCE.xlsx` | Section-wise RM length tolerances & raw size reference |
| `Grade details of ex grade and raw grade.xlsx` | External grade to raw grade compatibility mapping |
| `LFA BATCH DETAILS.xlsx` | Batch sizes, section weights, LFA batch lengths & billet weights |
| `CAMP PLAN FORMAT.xlsm` | Macro-enabled campaign output template |
| `M-1 Loss with Section.xlsx` | Root Radius (RR), thickness, and corrected section weight rules |
| `SPM Presentation .ppt` | Rail Mill process, capacity, and section technical parameters |
| `all sections.zip` / `data/section_workbooks/` | Extracted historical section planning workbooks |

---

## Project Structure

```
Rail_Mill_optimizer/
├── casting_plan/                     # Django Application
│   ├── management/commands/
│   │   └── import_spm_data.py        # Import command for reference workbooks
│   ├── templates/casting_plan/       # HTML5 Templates
│   ├── import_m1.py                  # M-1 Root Radius data importer
│   ├── models.py                     # Database schemas & models
│   ├── services.py                   # Casting plan optimization logic
│   ├── btr_service.py                # BTR ingestion & campaign generator
│   ├── rolling_plan_services.py      # Rolling plan generation engine
│   └── views.py                      # Request handlers & API endpoints
├── data/
│   ├── reference/                    # Plant reference workbooks & templates
│   └── section_workbooks/            # Extracted historical workbooks
├── media/                            # Generated campaign & rolling plan workbooks
├── rail_casting_plan/                # Django Project Configuration (settings, urls)
├── manage.py                         # Django CLI management script
├── README.md                         # Main documentation
├── requirements.txt                  # Python dependencies
└── run.bat                           # One-click Windows runner script
```

---

## Production & Plant Deployment Notes

> [!IMPORTANT]
> The optimizer and calculation engines use plant-provided reference heuristics (tolerances, Grade maps, LFA rules, 7% RM allowance, 87m cooling bed max). Active heat sizes, live ladle chemistry constraints, real-time mill prices, and live furnace scheduling parameters should be configured in **Master Data** or synced via SAP before automating final production dispatch.
