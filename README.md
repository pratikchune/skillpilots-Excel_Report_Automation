# Excel Report Automation

Place `.xlsx` or `.xlsm` files in `input/`.
Install dependencies with `python -m pip install -r requirements.txt`.
Run `python automation.py` (default group column: Department).
Choose another category using `python automation.py --group-by Region`.

The output folder receives a timestamped workbook with Cleaned_Data,
Numeric_Summary, Aggregated_Data, and File_Summary sheets, plus automation.log.
