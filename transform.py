#!/usr/bin/env python3
"""
transform.py — converts the daily "lastminute_com_group_cc_post_call_responses_*.xlsx"
export into dashboard_data.json for the Contact Centre KPI Explorer.

Usage:
    python3 transform.py input.xlsx dashboard_data.json

Designed to be run by the GitHub Action on a schedule (see
.github/workflows/update-data.yml), but works identically run by hand.
"""
import sys
import json
import csv
import openpyxl
import pandas as pd


def xlsx_to_dataframe(xlsx_path):
    """Stream the source workbook to a DataFrame without loading it all into
    memory at once (the export can be 90MB+). Auto-detects which row holds
    the real header (row 1 for a plain Google Sheet, row 3 for the original
    export-tool format with two title rows above it) by scanning the first
    few rows for the 'Local Response Date' column name."""
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb.worksheets[0]  # use whichever tab is first, regardless of its name
    rows_iter = ws.iter_rows(values_only=True)

    all_rows = []
    header_row_idx = None
    for i, row in enumerate(rows_iter):
        all_rows.append(row)
        if row and 'Local Response Date' in row:
            header_row_idx = i
        if i >= 4 and header_row_idx is not None:
            break  # found it, no need to keep buffering everything in memory

    if header_row_idx is None:
        raise RuntimeError(
            "Could not find a 'Local Response Date' column in the first few rows "
            "of the sheet. Check the export's column headers haven't changed."
        )

    tmp_csv = xlsx_path + '.tmp.csv'
    with open(tmp_csv, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(all_rows[header_row_idx])
        for row in all_rows[header_row_idx + 1:]:
            writer.writerow(row)
        for row in rows_iter:  # continue streaming the rest of the file
            writer.writerow(row)

    df = pd.read_csv(tmp_csv, low_memory=False)
    return df


DIM_MAP = {
    'Unit Name': 'Agent', 'Age Group': 'Age Group', 'Chat Language': 'Language',
    'Chat Entry Point': 'Entry Point', 'Chat Queue': 'Queue', 'Chat Closed By': 'Closed By',
    'Contact Reason Level1': 'Contact Reason L1', 'Contact Reason Level2': 'Contact Reason L2',
    'Contact Reason Level3': 'Contact Reason L3', 'Brand': 'Brand', 'Chat BP Country': 'Country',
    'Chat Product': 'Product', 'Flight Type.1': 'Flight Type', 'Booking Cluster': 'Booking Cluster',
    'Plus Flag': 'Plus Flag', 'Fare Type': 'Fare Type', 'Booking Category.1': 'Booking Category',
    'Insurance Flag.1': 'Insurance Flag', 'Car Flag.1': 'Car Flag', 'Booking Status.1': 'Booking Status',
    'Refund Flag.1': 'Refund Flag', 'Loyalty Flag': 'Loyalty Flag', 'Chat transferred': 'Chat Transferred',
    'Chat hold usage': 'Hold Usage', 'E2e Flag': 'E2E Flag', 'Issue Queue': 'Issue Queue',
    'Id_Source_Engine': 'Source Engine',
}

QA_MAP = {
    'Chat Soft Skills - Active listening - Empathy (12)': 'QA: Empathy',
    'Chat Soft Skills - Active listening - Ownership (8)': 'QA: Ownership',
    "Chat Soft Skills - Active listening - Asking the Customer to Repeat Information (8)": 'QA: No Repeat Requests',
    'Chat Soft Skills - Language - Courtesy & Positive Language (9)': 'QA: Courtesy & Positive Language',
    'Chat Soft Skills - Language - Unprofessional Language (14)': 'QA: No Unprofessional Language',
    'Chat Gathering Information - Acknowledging (12)': 'QA: Acknowledging Customer',
    'Chat Gathering Information - DPA Booking ID (0)': 'QA: DPA Booking ID',
    "Chat Gathering Information - DPA Customer's name (4)": 'QA: DPA Customer Name',
    'Chat Gathering Information - DPA Email address (0)': 'QA: DPA Email',
    'Chat Negative emotions (12)': 'QA: No Negative Emotions',
    'Chat Branding - Greeting (3)': 'QA: Greeting',
    'Chat Branding - Uncertainty (3)': 'QA: No Uncertainty',
    'Chat Branding - Additional Assistance/Follow-Up (4)': 'QA: Additional Assistance',
    'Chat Branding - Promote Self-Service (3)': 'QA: Promote Self-Service',
    'Chat Branding - Closing (3)': 'QA: Closing',
}

EFFORT_MAP = {
    'Chat Interaction Effort Score - Hard Effort - Channel Switching (6)': 'Effort: No Channel Switching',
    'Chat Interaction Effort Score - Hard Effort - Emotions (5)': 'Effort: No Negative Emotions',
    'Chat Interaction Effort Score - Hard Effort - Escalation Agent Repetition of Information Agent (6)': 'Effort: No Agent Repetition',
    'Chat Interaction Effort Score - Hard Effort - Escalation Customer (5)': 'Effort: No Customer Escalation',
    'Chat Interaction Effort Score - Hard Effort - Length of Time (6)': 'Effort: Good Length of Time',
    'Chat Interaction Effort Score - Hard Effort - Perception of Hard Effort (5)': 'Effort: Low Perceived Effort',
    'Chat Interaction Effort Score - Hard Effort - Repeat Interaction/Callback (6)': 'Effort: No Repeat Interaction',
    'Chat Interaction Effort Score - Hard Effort - Repetition of Information Customer (6)': 'Effort: No Customer Repetition',
    'Chat Interaction Effort Score - Hard Effort - Self-Service (5)': 'Effort: Self-Service Not Needed',
    'Chat Interaction Effort Score - Easy Effort - Agent Confirming Ease (5)': 'Effort: Agent Confirmed Ease',
    'Chat Interaction Effort Score - Easy Effort - Ease of Use (10)': 'Effort: Ease of Use',
    'Chat Interaction Effort Score - Easy Effort - Emotions (10)': 'Effort: Positive Emotions',
    'Chat Interaction Effort Score - Easy Effort - First Call Resolution (15)': 'Effort: First Contact Resolution',
    'Chat Interaction Effort Score - Easy Effort - Positive Ending (10)': 'Effort: Positive Ending',
}


def build_dataset(df):
    out = pd.DataFrame()
    d = pd.to_datetime(df['Local Response Date'])
    out['Date'] = d.dt.strftime('%Y-%m-%d')
    out['Month'] = d.dt.strftime('%Y-%m')
    out['Weekday'] = d.dt.day_name()

    for src, dst in DIM_MAP.items():
        out[dst] = df[src].fillna('N/A').astype(str)

    out['CSAT'] = df['Overall CSAT']
    # Standard NPS formula: classify each respondent as Promoter (+1, score 9-10),
    # Passive (0, score 7-8), or Detractor (-1, score 0-6). The mean of this
    # column * 100 equals "% Promoters - % Detractors" over all respondents
    # (passives count in the denominator, contribute 0 to the numerator) --
    # exactly the standard NPS calculation, computed per filtered slice in the UI.
    def _nps_class(v):
        if pd.isna(v):
            return pd.NA
        if v >= 9:
            return 1
        if v <= 6:
            return -1
        return 0
    out['NPS'] = df['Overall NPS'].apply(_nps_class).astype('Int64')
    out['Chat Duration (min)'] = df['Chat Duration (Minutes)'].round(2)
    out['Hold Duration (min)'] = df['Chat hold duration (minutes)'].round(2)
    out['QM Score'] = df['Chat - Agent QM Score']
    out['Interaction Effort Score'] = df['Chat Interaction Effort Score']

    for src, dst in QA_MAP.items():
        out[dst] = (df[src] == 2).astype('Int64')
        out.loc[df[src].isna(), dst] = pd.NA

    for src, dst in EFFORT_MAP.items():
        out[dst] = (df[src] == 2).astype('Int64')
        out.loc[df[src].isna(), dst] = pd.NA

    out['CSI More Details'] = df['CSI More Details']
    out['Url Chat Id'] = df['Url Chat Id']
    return out


def to_dashboard_json(out):
    dim_cols = list(DIM_MAP.values()) + ['Date', 'Month', 'Weekday']
    metric_cols = [c for c in out.columns if c not in dim_cols and c not in ('CSI More Details', 'Url Chat Id')]

    data = {'dims': {}, 'metrics': {}, 'n': len(out)}

    for c in dim_cols:
        codes, uniques = pd.factorize(out[c])
        data['dims'][c] = {'codes': codes.tolist(), 'cats': uniques.tolist()}

    for c in metric_cols:
        s = out[c]
        data['metrics'][c] = [None if pd.isna(v) else round(float(v), 3) for v in s]

    csat_int = out['CSAT'].apply(lambda v: 'N/A' if pd.isna(v) else str(int(round(v))))
    codes, uniques = pd.factorize(csat_int)
    order = ['5', '4', '3', '2', '1', 'N/A']
    remap = {i: order.index(cat) for i, cat in enumerate(uniques) if cat in order}
    data['dims']['CSAT'] = {'codes': [remap[c] for c in codes], 'cats': order}

    data['text'] = {
        'CSI More Details': [None if pd.isna(v) else str(v) for v in out['CSI More Details']],
        'Url Chat Id': [None if pd.isna(v) else str(v) for v in out['Url Chat Id']],
    }
    return data


def main():
    if len(sys.argv) != 3:
        print("Usage: python3 transform.py <input.xlsx> <output dashboard_data.json>")
        sys.exit(1)
    xlsx_path, out_path = sys.argv[1], sys.argv[2]

    print(f"Reading {xlsx_path} ...")
    df = xlsx_to_dataframe(xlsx_path)
    print(f"  {len(df):,} rows, {len(df.columns)} columns")

    print("Transforming ...")
    out = build_dataset(df)
    data = to_dashboard_json(out)

    print(f"Writing {out_path} ...")
    with open(out_path, 'w') as f:
        json.dump(data, f, separators=(',', ':'))
    print(f"Done: {data['n']:,} responses.")


if __name__ == '__main__':
    main()
