# -*- coding: utf-8 -*-
"""Timetable.ipynb
"""

import pandas as pd
import matplotlib.pyplot as plt

# Read our newly generated, clean, and highly variable timetable from Excel
df = pd.read_excel("timetable.xlsx", engine="openpyxl")

# Ensure proper date formatting if DATE column exists
if "DATE" in df.columns:
    df["DATE"] = pd.to_datetime(df["DATE"])
    num_days_in_data = df["DATE"].nunique()
else:
    num_days_in_data = 5 # default fallback

# Basic calculations
total_classes = len(df)
total_teaching_hours = df["DURATION"].sum()
total_classrooms = df["CLASSROOM"].nunique()
total_faculty = df["FACULTY"].nunique()

# Determine realistic available slots per room based on unique dates in dataset
slots_per_day = 8 # Standard 8 teaching slots/hours per day
available_hours = total_classrooms * num_days_in_data * slots_per_day

# Utilization calculations
utilization = (total_teaching_hours / available_hours) * 100
vacant_hours = max(0, available_hours - total_teaching_hours)

# Classroom-wise utilization scaled by unique dates in dataset
classroom_hours = df.groupby("CLASSROOM")["DURATION"].sum()
classroom_utilization = (
    classroom_hours / (num_days_in_data * slots_per_day)
) * 100

# Daily utilization
daily_hours = df.groupby("DAY")["DURATION"].sum()

# Create dashboard
fig, ax = plt.subplots(2, 2, figsize=(14, 9))

# 1. Classroom-wise utilization using a single consistent clean blue shade
ax[0, 0].bar(
    classroom_utilization.index,
    classroom_utilization.values,
    color="#3182bd",
    edgecolor="black"
)
ax[0, 0].set_title("Classroom-wise Utilization (%)")
ax[0, 0].set_xlabel("Classroom")
ax[0, 0].set_ylabel("Utilization (%)")
ax[0, 0].axhline(100, color="red", linestyle="--", label="100% Capacity")
ax[0, 0].legend()

# 2. Daily utilization trend showing significant workload shifts
days_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
daily_hours = daily_hours.reindex(days_order).fillna(0)

ax[0, 1].plot(
    daily_hours.index,
    daily_hours.values,
    marker="o",
    linewidth=2.5,
    color="forestgreen"
)
ax[0, 1].set_title("Teaching Load by Weekday (Total Hours)")
ax[0, 1].set_xlabel("Day")
ax[0, 1].set_ylabel("Total Hours")
ax[0, 1].tick_params(axis="x", rotation=30)

# 3. Occupied vs Vacant hours using a contrasting warm amber/orange palette to avoid dullness
ax[1, 0].pie(
    [total_teaching_hours, vacant_hours],
    labels=["Occupied Hours", "Vacant Hours"],
    autopct="%1.1f%%",
    startangle=90,
    colors=["#e6550d", "#f5f5f5"],
    wedgeprops={"edgecolor": "gray"}
)
ax[1, 0].set_title("Overall Capacity Allocation")

# 4. Summary box
ax[1, 1].axis("off")

summary = f"""
COLLEGE TIMETABLE ANALYSIS

Total Classes Scheduled: {total_classes}
Total Classrooms Active: {total_classrooms}
Total Faculty Members: {total_faculty}

Total Teaching Hours: {total_teaching_hours} hrs
Total Capacity ({num_days_in_data} Days @ {slots_per_day} hrs/day): {available_hours} hrs
Unused/Vacant Hours: {vacant_hours} hrs

Overall Timetable Utilization: {utilization:.2f}%
"""

ax[1, 1].text(
    0.05, 0.5,
    summary,
    fontsize=11,
    verticalalignment="center",
    fontfamily="monospace"
)

# Dashboard title
plt.suptitle(
    "College Timetable Utilization & Analytical Dashboard",
    fontsize=18,
    fontweight="bold"
)

plt.tight_layout()

# Save the dashboard image to the workspace
plt.savefig("timetable_dashboard.png", dpi=300, bbox_inches="tight")
print("Dashboard image saved successfully as 'timetable_dashboard.png'.")

plt.show()

import pandas as pd
import numpy as np

# 1. Load the original dataset
df_raw = pd.read_excel("timetable.xlsx", engine="openpyxl")

# 2. Make columns and values cleaner and consistent
df_clean = df_raw.copy()

# Ensure clean column labels
df_clean.columns = [col.strip().upper() for col in df_clean.columns]

# Standardize text columns (strip whitespace, title case for names)
text_cols = ["DEPARTMENT", "FACULTY", "SUBJECT", "CLASSROOM", "DAY"]
for col in text_cols:
    if col in df_clean.columns:
        df_clean[col] = df_clean[col].astype(str).str.strip().str.title()

# 3. Sort chronologically if Date or Day exists
day_mapping = {
    "Monday": 1,
    "Tuesday": 2,
    "Wednesday": 3,
    "Thursday": 4,
    "Friday": 5,
    "Saturday": 6,
    "Sunday": 7
}

if "DAY" in df_clean.columns:
    df_clean["DAY_NUM"] = df_clean["DAY"].map(day_mapping)

# Define sorting priorities
sort_cols = []
if "DATE" in df_clean.columns:
    df_clean["DATE"] = pd.to_datetime(df_clean["DATE"])
    sort_cols.append("DATE")
elif "DAY_NUM" in df_clean.columns:
    sort_cols.append("DAY_NUM")

if "TIME" in df_clean.columns:
    sort_cols.append("TIME")
elif "TIME_SLOT" in df_clean.columns:
    sort_cols.append("TIME_SLOT")

if sort_cols:
    df_clean = df_clean.sort_values(by=sort_cols).reset_index(drop=True)

# Drop temporary helper column
if "DAY_NUM" in df_clean.columns:
    df_clean = df_clean.drop(columns=["DAY_NUM"])

# 4. Save the cleaned, highly readable file back to the workspace
output_filename = "convenient_college_timetable.xlsx"
df_clean.to_excel(output_filename, index=False, engine="openpyxl")

print(f"Success! Your data has been cleaned and organized chronologically.")
print(f"Saved as: '{output_filename}' in your sidebar.")
display(df_clean.head(10))
