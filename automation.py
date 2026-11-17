"""Reusable Excel report cleaning and aggregation tool."""
from __future__ import annotations
import argparse
import logging
from datetime import datetime
from pathlib import Path
from typing import Optional
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
LOG_PATH = OUTPUT_DIR / "automation.log"


def setup_logging() -> None:
    """Set up logging to the console and output/automation.log."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s: %(message)s",
        handlers=[logging.FileHandler(LOG_PATH, encoding="utf-8"),
                  logging.StreamHandler()],
        force=True,
    )


def clean_excel_data(df: pd.DataFrame) -> pd.DataFrame:
    """Remove empty rows/columns, trim text, and remove duplicate rows."""
    df = df.dropna(how="all").dropna(axis=1, how="all").copy()
    df.columns = [str(c).strip() for c in df.columns]
    for col in df.select_dtypes(include=["object", "string"]).columns:
        df[col] = df[col].astype("string").str.strip().replace("", pd.NA)
    return df.drop_duplicates()


def read_excel_reports(input_dir: Path = INPUT_DIR) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Read worksheets sequentially and return combined data and file counts."""
    if not input_dir.is_dir():
        raise FileNotFoundError(f"Create an input folder at: {input_dir}")
    files = sorted(p for pattern in ("*.xlsx", "*.xlsm")
                   for p in input_dir.glob(pattern)
                   if not p.name.startswith("~$"))
    if not files:
        raise FileNotFoundError(f"No Excel files found in {input_dir}")

    frames = []
    summary_rows = []
    for path in files:
        count = 0
        logging.info("Reading %s", path.name)
        try:
            with pd.ExcelFile(path) as book:
                for sheet in book.sheet_names:
                    frame = clean_excel_data(pd.read_excel(book, sheet_name=sheet))
                    if frame.empty:
                        continue
                    count += len(frame)
                    frame["Source_File"] = path.name
                    frame["Source_Sheet"] = sheet
                    frames.append(frame)
        except Exception:
            logging.exception("Could not read %s", path.name)
        summary_rows.append({"Source_File": path.name, "Total_Records": count})

    if not frames:
        raise ValueError("No usable worksheet data was found.")
    return pd.concat(frames, ignore_index=True, sort=False), pd.DataFrame(summary_rows)


def build_report_sheets(df: pd.DataFrame, file_summary: pd.DataFrame,
                        group_by: str) -> dict[str, pd.DataFrame]:
    """Create report DataFrames separately from writing them to Excel."""
    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    if numeric_cols:
        numeric = df[numeric_cols].agg(["count", "sum", "mean", "min", "max"]).T.reset_index()
        numeric.columns = ["Column", "Count", "Total", "Average", "Minimum", "Maximum"]
    else:
        numeric = pd.DataFrame(columns=["Column", "Count", "Total", "Average", "Minimum", "Maximum"])

    if group_by in df.columns:
        value_cols = [c for c in numeric_cols if c != group_by]
        if value_cols:
            grouped = df.groupby(group_by, dropna=False)[value_cols].agg(["sum", "mean", "count"])
            grouped.columns = [f"{c}_{stat}" for c, stat in grouped.columns]
            grouped = grouped.reset_index()
        else:
            grouped = pd.DataFrame({"Message": ["No numeric columns available for grouping."]})
    else:
        grouped = pd.DataFrame({"Message": [
            f"Grouping column '{group_by}' not found.",
            "Available columns: " + ", ".join(map(str, df.columns))
        ]})

    return {"Cleaned_Data": df, "Numeric_Summary": numeric,
            "Aggregated_Data": grouped, "File_Summary": file_summary}


def write_report(sheets: dict[str, pd.DataFrame], output_dir: Path = OUTPUT_DIR) -> Path:
    """Write the prepared report sheets to a timestamped Excel workbook."""
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / f"Final_Report_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name, index=False)
    logging.info("Saved report: %s", path)
    return path


def generate_reports(df: pd.DataFrame, file_summary: Optional[pd.DataFrame] = None,
                     group_by: str = "Department") -> Path:
    """Build and write all output worksheets."""
    if file_summary is None:
        file_summary = pd.DataFrame([{"Source_File": "Combined input",
                                      "Total_Records": len(df)}])
    return write_report(build_report_sheets(df, file_summary, group_by))


def main() -> int:
    """Parse command-line options and execute the automation."""
    parser = argparse.ArgumentParser(description="Clean and aggregate Excel reports.")
    parser.add_argument("--group-by", default="Department",
                        help="Column for grouped aggregation (default: Department).")
    args = parser.parse_args()
    setup_logging()
    logging.info("Starting Excel report automation.")
    try:
        data, file_summary = read_excel_reports()
        report = generate_reports(data, file_summary, args.group_by)
        logging.info("Completed successfully: %s", report)
        return 0
    except Exception as exc:
        logging.exception("Automation failed: %s", exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
