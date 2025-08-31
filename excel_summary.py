# excel_summary.py
import sys
import pandas as pd

def main(src='input.xlsx', dst='report.xlsx', sheet='Data', category_col='Category', value_col='Value'):
    # Read
    df = pd.read_excel(src, sheet_name=sheet)

    # Validate columns
    for col in (category_col, value_col):
        if col not in df.columns:
            raise ValueError(f"Column '{col}' not found. Available columns: {list(df.columns)}")

    # Build summary
    summary = (
        df.groupby(category_col, dropna=False)[value_col]
          .agg(Count='count', Sum='sum', Mean='mean')
          .reset_index()
          .sort_values('Sum', ascending=False)
    )

    # Write
    with pd.ExcelWriter(dst, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Raw')
        summary.to_excel(writer, index=False, sheet_name='Summary')

    print(f"OK: wrote {dst}")

if __name__ == '__main__':
    # Allow CLI overrides
    args = sys.argv[1:]
    main(*args)  # src, dst, sheet, category_col, value_col
