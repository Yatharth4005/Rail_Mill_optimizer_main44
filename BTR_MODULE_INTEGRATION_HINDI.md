# Upload BTR Module — Integration Guide

## Updated ZIP ko direct run karna

1. ZIP ko extract karein.
2. Project folder me terminal kholein.
3. Windows par `run.bat` run karein, ya manually:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python manage.py migrate
python manage.py import_spm_data
python manage.py runserver 127.0.0.1:8000
```

4. Browser me `http://127.0.0.1:8000/upload-btr/` open karein.

`python manage.py migrate` zaroor chalana hai, kyunki migration `0006_btrcampaign.py` campaign history ke liye nayi database table banati hai.

## Sirf module ko existing portal me merge karna

In files ko existing project me copy/merge karein:

| File | Kaam |
|---|---|
| `casting_plan/btr_services.py` | BTR reading, calculation, sorting aur CAMP Excel generation |
| `casting_plan/templates/casting_plan/upload_btr.html` | Upload BTR page ka UI |
| `casting_plan/migrations/0006_btrcampaign.py` | Campaign history database table |
| `casting_plan/test_btr.py` | Sample BTR generation tests |
| `data/reference/CAMP PLAN FORMAT.xlsm` | Macro-enabled output template |
| `data/reference/BTR_file_sample.xlsx` | QA/sample input |

In existing files ke changes merge karein:

| File | Required change |
|---|---|
| `casting_plan/models.py` | `BTRCampaign` model |
| `casting_plan/views.py` | Upload, generate, download aur delete views |
| `spm_casting_plan/urls.py` | `/upload-btr/` URLs |
| `casting_plan/templates/casting_plan/base.html` | Sidebar ka **Upload BTR** link |

Merge ke baad:

```powershell
pip install -r requirements.txt
python manage.py migrate
python manage.py check
python manage.py runserver 127.0.0.1:8000
```

## Portal me actual flow

1. User Upload BTR page par Excel select karta hai.
2. Server mandatory headers validate karta hai aur Planning Material list banata hai.
3. User ek Planning Material choose karta hai.
4. System sirf us material ki positive `Bal Roll Qty` rows leta hai.
5. Zero Length rows me valid Min/Max hone par Max Length use hoti hai; invalid rows skip hoti hain.
6. `Wt./ Pc` aur `No. of Pcs.` calculate hote hain.
7. HT, MS aur PANTHER rows ko alag group nahi kiya jata; sab records thickness-wise saath sort hote hain.
8. Har thickness ka ek combined subtotal aur akhir me grand total add hota hai.
9. Sheet1 me final campaign aur Sheet2 me source audit data write hota hai.
10. `.xlsm` file Campaign History me download ke liye milti hai; template ke VBA macros preserve rehte hain.

## Sorting sequence

| Priority | Field | Direction |
|---:|---|---|
| 1 | Thickness | Small to large |
| 2 | Cutting Length | Large to small |
| 3 | No. of Pcs. | High to low |
| 4 | Embossing | A to Z |

HT/MS/PANTHER category audit summary me available rehti hai, lekin Sheet1 sorting aur subtotal ke samay categories alag nahi hoti hain.

## Important folders

- Uploaded BTR: `media/btr_uploads/`
- Generated campaign files: `media/btr_campaigns/`
- CAMP template: `data/reference/CAMP PLAN FORMAT.xlsm`

Production deployment me `media/` folder ko application process ke liye writable rakhein aur regular backup/cleanup policy configure karein.
