from datetime import datetime, timedelta
import pandas as pd
import numpy as np
import argparse
import xlrd
import xlwt
import matplotlib.pyplot as plt

parser = argparse.ArgumentParser()
parser.add_argument(
    "--input", required=True, help="Path to Summary Report file you wish to analyse"
)
parser.add_argument(
    "--output-excel", default="output.xls", help="Name of excel file to output modified report"
)
parser.add_argument(
    "--output-html", default="output.html", help="Name of html file to output modified report"
)
opts = parser.parse_args()

# pd.read_excel not working on windows machine. We will manually read in the file
book = xlrd.open_workbook(opts.input)
sh = book.sheet_by_index(0)
data = np.array([sh.row_values(i) for i in range(sh.nrows)])
data_dict = {col: data[:,num][1:] for num, col in enumerate(data[0])}
df = pd.DataFrame.from_dict(data_dict)
# continue with rest of script
unique_names = np.unique(df["Pupil Name"])
initial_entry = np.zeros_like(unique_names)
new_df = pd.DataFrame.from_dict(
    {
        "Student Name": unique_names,
        "Number of Behaviour Rewards": initial_entry,
        "Number of Academic Rewards": initial_entry,
        "Number of Academic Sanctions": initial_entry,
        "Number of Behaviour Sanctions": initial_entry,
        "Total Value": initial_entry
    }
)

for num, name in enumerate(unique_names):
    bool = df["Pupil Name"] == name
    _type = df["Type of Sanction"][bool]
    new_df["Number of Behaviour Rewards"][num] = np.sum(_type == "Behaviour Reward")
    new_df["Number of Academic Rewards"][num] = np.sum(_type == "Academic Reward")
    new_df["Number of Behaviour Sanctions"][num] = np.sum(_type == "Behaviour Sanction")
    new_df["Number of Academic Sanctions"][num] = np.sum(_type == "Academic Sanction")
    new_df["Total Value"][num] = np.sum(df["Points value of sanction"][bool].astype(np.float32))

# new_df.to_excel not working on windows machine. We will manually save to excel file
wb = xlwt.Workbook()
ws = wb.add_sheet('Sheet1')
for idx, (col, value) in enumerate(new_df.items()):
    ws.write(0, idx, col)
    for num in range(len(value)):
        ws.write(num + 1, idx, str(value[num]))
wb.save(opts.output_excel)
new_df.to_html(buf=opts.output_html, index=False)

# Now we can do some statistics
dates = np.array([datetime.strptime(_, "%d/%m/%Y") for _ in df["Date of sanction"]])
first = np.min(dates)
weeks = {}
num = 1
while True:
    end = first + timedelta(days=7 - first.weekday())
    bool = (dates >= first) & (dates < end)
    mask = lambda s: df["Pupil Name"] == s
    weeks[num] = {
        student: {
            "Number of Sanctions": np.sum(["Sanction" in _ for _ in df["Type of Sanction"][bool][mask(student)]]),
            "Number of Rewards": np.sum(["Reward" in _ for _ in df["Type of Sanction"][bool][mask(student)]]),
            "Total Value": np.sum(df["Points value of sanction"][bool][mask(student)].astype(np.float32))
        } for student in unique_names
    }
    if first > np.max(dates):
        break
    else:
        first += timedelta(days=7 - first.weekday())
        num += 1

fig, axs = plt.subplots(figsize=(20, 12), nrows=3, sharex=True)
nweeks = len(weeks.keys())
width = 0.95 / nweeks
for num in range(len(unique_names)):
    multiplier = 1
    for idx in weeks.keys():
        if num == 0:
            label = f"Week {idx}"
        else:
            label = None
        offset = width * multiplier
        rects = axs[0].bar(num - (width * nweeks / 2) + offset, weeks[idx][unique_names[num]]["Total Value"], width, label=label, color=f"C0{idx}")
        rects = axs[1].bar(num - (width * nweeks / 2) + offset, weeks[idx][unique_names[num]]["Number of Rewards"], width, label=label, color=f"C0{idx}")
        rects = axs[2].bar(num - (width * nweeks / 2) + offset, weeks[idx][unique_names[num]]["Number of Sanctions"], width, label=label, color=f"C0{idx}")
        multiplier += 1
    axs[0].axvline(num - width * nweeks / 2, color='lightgrey', linestyle=":")
    axs[1].axvline(num - width * nweeks / 2, color='lightgrey', linestyle=":")
    axs[2].axvline(num - width * nweeks / 2, color='lightgrey', linestyle=":")
axs[0].set_xticklabels("")
axs[1].set_xticklabels("")
axs[0].legend(loc="upper center", bbox_to_anchor=(0.5, 1.1), ncol=nweeks)
axs[1].set_xticks(np.arange(len(unique_names)))
axs[2].set_xticklabels(unique_names, rotation=45)
axs[0].set_ylabel("Weekly total")
axs[1].set_ylabel("Total rewards")
axs[2].set_ylabel("Total sanctions")
plt.tight_layout()
plt.savefig("week_by_week.png")
plt.close()

raw = open("output.html", "r")
lines = raw.readlines()
with open("output.html", "w") as f:
    lines += ["<p></p>\n"]
    lines += ["<p></p>\n"]
    lines += ["<img src='week_by_week.png' style='width:85%'></img>\n"]
    f.writelines(lines)
