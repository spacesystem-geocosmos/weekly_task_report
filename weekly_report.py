import pandas as pd
from datetime import datetime, timedelta
import os
import sys

# change working directory to the script/exe's directory so output goes there
os.chdir(os.path.dirname(os.path.abspath(sys.argv[0])))

# ===== load the workbook =====
file = "Space Systems Division.xlsx"   # adjust if the filename differs
df = pd.read_excel(file)

# ===== clean & filter =====
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
# use only the date portion for comparisons so "due today" isn’t marked late
today_date = today.date()
week_ago = today - timedelta(days=7)

# parse dates
# keep full timestamps but we'll compare their .date() below
df["Due date"] = pd.to_datetime(df["Due date"], errors="coerce")
df["Completed Date"] = pd.to_datetime(df["Completed Date"], errors="coerce")

# only consider items that aren't in the Completed bucket and aren't already marked
# as Completed in the Progress column (some exports mix the two) for most metrics
open_df = df[(df["Bucket Name"] != "Completed") & (df["Progress"] != "Completed")]

# convenience for counting on either df or open_df

def count_by(condition, use_open=True):
    data = open_df if use_open else df
    return data[condition].groupby("_assignee").size()

# late means strictly before today’s date
late_tasks_count = count_by((open_df["Due date"].dt.date < today_date) &
                            (open_df["Bucket Name"] == "Tasks"))
# RFR items should also be open
rfr_late_count = count_by((open_df["Due date"].dt.date < today_date) &
                          (open_df["Bucket Name"] == "Ready For Review"))
open_count = count_by(open_df["Bucket Name"] != "Completed")  # trivially open_df
urgent_count = count_by(open_df["Priority"].str.contains("Urgent",
                                                         case=False, na=False))
late_ratio = (late_tasks_count / open_count).fillna(0)

# weekly lists
done_last_week = df[(df["Bucket Name"] == "Completed") &
                    (df["Completed Date"] >= week_ago)]
planned_next_week = df[(df["Bucket Name"] != "Completed") &
                       (df["Due date"] >= today) &
                       (df["Due date"] <= today + timedelta(days=7))]

# build & print snapshot table
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

# build the late-details table (flat, useful for debugging or export)
late_details = open_df[
    (open_df["Due date"].dt.date < today_date)
][["_assignee", "Task Name", "Bucket Name", "Due date"]].rename(columns={"_assignee": "Assignee"})

# convert the weekly lists to dataframes as well

done_df = pd.DataFrame({"Task Name": done_last_week["Task Name"].unique()})
planned_df = pd.DataFrame({"Task Name": planned_next_week["Task Name"].unique()})

# export everything to an Excel workbook
import os
print("current working directory is", os.getcwd())
output_file = f"weekly_report_output_{today.strftime('%Y-%m-%d')}.xlsx"
print("will write to", os.path.abspath(output_file))
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    metrics.to_excel(writer, sheet_name="Metrics")
    late_details.to_excel(writer, sheet_name="Late Details", index=False)
    done_df.to_excel(writer, sheet_name="Done Last Week", index=False)
    planned_df.to_excel(writer, sheet_name="Planned Next Week", index=False)

print(f"Wrote report to {output_file}")

print("\n===== DONE LAST WEEK =====")
for task in done_last_week["Task Name"].unique():
    print("-", task)

print("\n===== PLANNED FOR NEXT WEEK =====")
for task in planned_next_week["Task Name"].unique():
    print("-", task)
