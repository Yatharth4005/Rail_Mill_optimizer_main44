# Rolling Plan Integration

## User flow

1. Upload BTR and generate a campaign as before.
2. In **Campaign & Rolling Plan History**, enter the rolling-plan date and optional camp number.
3. Click **Generate Rolling Plan**.
4. Review the calculated blooms, RM quantity and yield shown in the history row.
5. Download the generated macro-enabled `.xlsm` workbook.

The output is a copy of the generated campaign workbook, so its VBA project and campaign sheets remain intact. It also includes:

- `Final Plan`
- `Grade Mapping`
- `Raw Material Mapping`
- `Plan Assumptions`

## Implemented rules

- Thickness, section weight, external grade and rolling quantity come from campaign `Sheet1`.
- Rolling quantity equals campaign `BTR`.
- Raw grade comes from `data/reference/Grade details of ex grade and raw grade.xlsx`.
- For angles, Rolling Section is Planning Mat + positive thickness + the BIS/BSEN or Remarks standard suffix; non-angle sections do not append thickness.
- Raw material size is resolved from the Product List. Sectional weights are BL_165X165=214, BL_250X250=490, BL_130X130=132 and BB_355X280X90=480 kg/m.
- R.M. Grade matches Ext Grade + STD. Billet sectional weight >300 kg/m uses Higher Section; <=300 kg/m uses Lighter Section.
- Sorting order is STD, raw material, embossing, then thickness for angle sections (`AL_...`) or section weight for every other section.
- After each contiguous Rolling Section block produced by that sorting order, one subtotal row sums only Rolling Qty and Bloom 200x200 (PCS).
- Bloom Cum continues across Rolling Section totals.
- Bloom pieces use cumulative BTR with the 7% RM allowance.
- RM quantity, bloom cumulative, rolled-cut pieces and rolled length are Excel formulas.
- Cooling-bed divisor begins at 2 and increases until the calculated bed length is strictly below 87 m.
- Cutting plan is provisionally calculated as `cutting length × whole pieces fitting within cooling-bed length`.
- Remaining order/customer fields come from campaign `Sheet1`.

## Provisional settings

Until final plant rules are supplied, these values are used and displayed in `Plan Assumptions`:

- Billet length: 11.65 m
- RM allowance: 7%
- Cooling-bed maximum: strictly below 87 m

## Deployment

Run the new database migration before starting the updated application:

```powershell
python manage.py migrate
```

Then start the application normally using `run.bat` or `python manage.py runserver`.
