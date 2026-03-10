import pandas as pd
from datetime import datetime, timedelta
import os
import sys

# change working directory to the script/exe's directory so output goes there
os.chdir(os.path.dirname(os.path.abspath(sys.argv[0])))

DATA_DIR = "data"
OUTPUT_DIR = "reports"
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ===== process all Excel files in the data directory =====
input_files = [f for f in os.listdir(DATA_DIR) if f.lower().endswith('.xlsx')]
if not input_files:
    print(f"No Excel (.xlsx) files found in the '{DATA_DIR}' directory.")
    sys.exit(1)

for file in input_files:
    file_path = os.path.join(DATA_DIR, file)
    base = os.path.splitext(os.path.basename(file))[0]
    output_file = os.path.join(OUTPUT_DIR, f"{base}_weekly_report_{datetime.today().strftime('%Y-%m-%d')}.xlsx")
    if os.path.exists(output_file):
        print(f"Skipping {file} (report already exists: {output_file})")
        continue
    print(f"Processing {file} ...")
    df = pd.read_excel(file_path, engine="openpyxl")
    df = df.fillna("")
    # drop buckets with blocked/archive/discontinued rows
    df = df[~df["Bucket Name"].str.contains(
        "Blocked|Archive|Archived|Discontinued", case=False, na=False)]
    # explode multi‑assignee cells so we can group by person
    assignees = df["Assigned To"].astype(str).str.split(";")
    df = df.assign(_assignee=assignees).explode("_assignee")
    df["_assignee"] = df["_assignee"].str.strip()
    # keep only the three people of interest
    team = ["Vachik Khachatryan", "Sargis Pinamyan", "Eliza Ayvazyan"]
    df = df[df["_assignee"].isin(team)]
    today = datetime.today()
    today_date = today.date()
    week_ago = today - timedelta(days=7)
    # parse dates
    df["Due date"] = pd.to_datetime(df["Due date"], errors="coerce")
    df["Completed Date"] = pd.to_datetime(df["Completed Date"], errors="coerce")
    open_df = df[(df["Bucket Name"] != "Completed") & (df["Progress"] != "Completed")]
    def count_by(condition, use_open=True):
        data = open_df if use_open else df
        return data[condition].groupby("_assignee").size()
    late_tasks_count = count_by((open_df["Due date"].dt.date < today_date) &
                                (open_df["Bucket Name"] == "Tasks"))
    rfr_late_count = count_by((open_df["Due date"].dt.date < today_date) &
                              (open_df["Bucket Name"] == "Ready For Review"))
    open_count = count_by(open_df["Bucket Name"] != "Completed")
    urgent_count = count_by(open_df["Priority"].str.contains("Urgent", case=False, na=False))
    late_ratio = (late_tasks_count / open_count).fillna(0)
    done_last_week = df[(df["Bucket Name"] == "Completed") &
                        (df["Completed Date"] >= week_ago)]
    planned_next_week = df[(df["Bucket Name"] != "Completed") &
                           (df["Due date"] >= today) &
                           (df["Due date"] <= today + timedelta(days=7))]
    date_str = today.strftime("%m/%d/%Y")
    metrics = pd.DataFrame(
        index=[
            "Tasks (Late)",
            "Ready For Review (Late)",
            "Total open",
            "Total urgent tasks",
            "Late task ratio",
        ]
    )
    for person in team:
        metrics[person] = [
            late_tasks_count.get(person, 0),
            rfr_late_count.get(person, 0),
            open_count.get(person, 0),
            urgent_count.get(person, 0),
            round(late_ratio.get(person, 0), 2),
        ]
    metrics["Total"] = metrics[team].sum(axis=1)
    if open_count.sum() > 0:
        metrics.at["Late task ratio", "Total"] = round(
            late_tasks_count.sum() / open_count.sum(), 2)
    else:
        metrics.at["Late task ratio", "Total"] = 0
    late_details = open_df[
        (open_df["Due date"].dt.date < today_date)
    ][["_assignee", "Task Name", "Bucket Name", "Due date"]].rename(columns={"_assignee": "Assignee"})
    done_df = pd.DataFrame({"Task Name": done_last_week["Task Name"].unique()})
    planned_df = pd.DataFrame({"Task Name": planned_next_week["Task Name"].unique()})
    print("Will write to", os.path.abspath(output_file))
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        metrics.to_excel(writer, sheet_name="Metrics")
        late_details.to_excel(writer, sheet_name="Late Details", index=False)
        done_df.to_excel(writer, sheet_name="Done Last Week", index=False)
        planned_df.to_excel(writer, sheet_name="Planned Next Week", index=False)
    print(f"Wrote report to {output_file}\n")
    print("===== DONE LAST WEEK =====")
    for task in done_last_week["Task Name"].unique():
        print("-", task)
    print("===== PLANNED FOR NEXT WEEK =====")
    for task in planned_next_week["Task Name"].unique():
        print("-", task)

