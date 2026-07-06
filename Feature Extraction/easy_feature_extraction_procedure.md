# Procedure: Easy Variable Extraction Automation

## Goal

Automate the easiest bushing variables first, then append them as one row into the assistive model dataset.

This first stage targets:

- Bushing ID
- Conductor diameter
- Conductor material
- Shield length
- Shield material
- Shell material
- BIL voltage
- Max E-field over known conductor/shield/shed distance ranges
- Max charge over known conductor/shield/shed distance ranges

## File Structure

Place these files in the same folder as your original pipeline:

```text
API_Automation.py
easy_feature_extraction.py
```

Optional reference file:

```text
API_Automation_with_easy_feature_integration.py
```

## Implementation Location

Inside your original `API_Automation.py`, insert the integration section after:

1. The ELECTRO project is opened
2. The simulation has run
3. E-field and/or charge CSV files have been exported

The correct location is usually near the end of the pipeline.

## Required Inputs

You need to provide the distance ranges from the exported ELECTRO line plot.

Example:

```python
conductor_range = (0.0, 25.0)
shield_range = (120.0, 170.0)
shed_range = (250.0, 500.0)
```

These are not universal. They depend on the distance coordinate from your ELECTRO exported plot.

## Output

The script appends one row to:

```text
outputs/extracted_feature_rows.csv
```

Each row can then be used as an ML feature row.

## Recommended First Test

Start by only testing output extraction from already-exported CSV files.

Use:

```python
from easy_feature_extraction import extract_max_outputs_from_distance_ranges

results = extract_max_outputs_from_distance_ranges(
    e_csv_path="exports/E_vs_distance.csv",
    charge_csv_path="exports/charge_vs_distance.csv",
    conductor_range=(0.0, 25.0),
    shield_range=(120.0, 170.0),
    shed_range=(250.0, 500.0),
)

print(results)
```

Once that works, connect the API-based metadata extraction.
