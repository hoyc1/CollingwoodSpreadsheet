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
unique_years = np.unique(df["Year"])
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

for year in unique_years:
    bool = year == df["Year"]
    average_value = np.mean(
        [np.sum(df["Points value of sanction"][bool][df["Pupil Name"] == student].astype(np.float32)) for student in unique_names]
    )
    print(f"Average number of points per student in {year}: {average_value}")

#I'm getting a bit obsessed with data...  As well as splitting the graphs into three, would it also be possible to graph the total rewards and sanctions (separately numbers and points value) of the whole House and each year group as a function of time?

# Now we can do some statistics
dates = np.array([datetime.strptime(_, "%d/%m/%Y") for _ in df["Date of sanction"]])
first = np.min(dates)
weeks = {}
num = 1
while True:
    end = first + timedelta(days=7 - first.weekday())
    bool = (dates >= first) & (dates < end)
    mask = lambda s: df["Pupil Name"] == s
    weeks[num] = {}
    for list, df_name in zip([unique_names, unique_years], ["Pupil Name", "Year"]):
        mask = lambda s: df[df_name] == s
        for v in list:
            value = df["Points value of sanction"][bool][mask(v)].astype(np.float32)
            sanction_mask = np.array(["Sanction" in _ for _ in df["Type of Sanction"][bool][mask(v)]])
            reward_mask = np.array(["Reward" in _ for _ in df["Type of Sanction"][bool][mask(v)]])
            weeks[num][v] = {
                "Number of Sanctions (Number)": np.sum(sanction_mask),
                "Number of Sanctions (Value)": np.sum(value[sanction_mask]),
                "Number of Rewards (Number)": np.sum(reward_mask),
                "Number of Rewards (Value)": np.sum(value[reward_mask]),
                "Total Value": np.sum(value)
            }
    if first > np.max(dates):
        break
    else:
        first += timedelta(days=7 - first.weekday())
        num += 1

plots = ["Total Value", "Number of Rewards", "Number of Sanctions"]
nweeks = len(weeks.keys())
width = 0.95 / nweeks
for _type in plots:
    if _type == "Total Value":
        fig = plt.figure(figsize=(5, 15))
        axs = [plt.gca()]
    else:
        fig, axs = plt.subplots(figsize=(10, 15), ncols=2, sharey=True)
    for num in range(len(unique_names)):
        multiplier = 1
        for idx in weeks.keys():
            if num == 0:
                label = f"Week {idx}"
            else:
                label = None
            offset = width * multiplier
            if _type == "Total Value":
                rects = axs[0].barh(num - (width * nweeks / 2) + offset, weeks[idx][unique_names[num]][f"{_type}"], width, label=label, color=f"C0{idx}")
                multiplier += 1
                continue
            rects = axs[0].barh(num - (width * nweeks / 2) + offset, weeks[idx][unique_names[num]][f"{_type} (Number)"], width, label=label, color=f"C0{idx}")
            rects = axs[1].barh(num - (width * nweeks / 2) + offset, weeks[idx][unique_names[num]][f"{_type} (Value)"], width, label=label, color=f"C0{idx}")
            multiplier += 1
        axs[0].axhline(num - width * nweeks / 2, color='lightgrey', linestyle=":")
        if _type != "Total Value":
            axs[1].axhline(num - width * nweeks / 2, color='lightgrey', linestyle=":")

    axs[0].set_yticks(np.arange(len(unique_names)))
    axs[0].set_yticklabels(unique_names)
    if _type == "Total Value":
        axs[0].set_xlabel("Total Value")
        axs[0].legend(loc="upper center", bbox_to_anchor=(0.5, 1.08), ncol=len(weeks) // 3)
        plt.tight_layout()
    else:
        plt.suptitle(_type)
        axs[0].legend(loc="upper center", bbox_to_anchor=(0.8, 1.07), ncol=len(weeks) // 2)
        handles, labels = axs[0].get_legend_handles_labels()
        fig.legend(handles, labels, loc='upper center', ncol=len(weeks) // 2, bbox_to_anchor=(0.5, 0.95))
        axs[0].get_legend().remove()
        axs[0].set_xlabel("Number")
        axs[1].set_xlabel("Value")
        xticks = axs[0].get_xticks()
        axs[0].set_xticks(np.arange(xticks[0], xticks[-1], 2).astype(int))
    plt.savefig(f"{_type.replace(' ', '_')}.png")
    plt.close()

fig, axs = plt.subplots(figsize=(10, 15), ncols=2, sharey=True)
for _type in plots[1:]:
    fig, axs = plt.subplots(figsize=(10, 15), ncols=2, sharey=True)
    for num in range(len(unique_years)):
        multiplier = 1
        for idx in weeks.keys():
            if num == 0:
                label = f"Week {idx}"
            else:
                label = None
            offset = width * multiplier
            rects = axs[0].barh(num - (width * nweeks / 2) + offset, weeks[idx][unique_years[num]][f"{_type} (Number)"], width, label=label, color=f"C0{idx}")
            rects = axs[1].barh(num - (width * nweeks / 2) + offset, weeks[idx][unique_years[num]][f"{_type} (Value)"], width, label=label, color=f"C0{idx}")
            multiplier += 1
        axs[0].axhline(num - width * nweeks / 2, color='lightgrey', linestyle=":")
        axs[1].axhline(num - width * nweeks / 2, color='lightgrey', linestyle=":")
    axs[0].set_yticks(np.arange(len(unique_years)))
    axs[0].set_yticklabels(unique_years)
    axs[0].legend(loc="upper center", bbox_to_anchor=(0.8, 1.07), ncol=len(weeks) // 2)
    handles, labels = axs[0].get_legend_handles_labels()
    fig.legend(handles, labels, loc='upper center', ncol=len(weeks) // 2, bbox_to_anchor=(0.5, 0.95))
    axs[0].get_legend().remove()
    axs[0].set_xlabel("Number")
    axs[1].set_xlabel("Value")
    xticks = axs[0].get_xticks()
    axs[0].set_xticks(np.arange(xticks[0], xticks[-1], 5).astype(int))
    plt.savefig(f"{_type.replace(' ', '_')}_year.png")
    plt.close()

raw = open("output.html", "r")
lines = raw.readlines()
with open("output.html", "w") as f:
    lines += ["<p></p>\n"]
    lines += ["<p></p>\n"]
    for _type in plots:
        if _type == "Total Value":
            lines += [f"<img src='{_type.replace(' ', '_')}.png' style='width:19%'></img>\n"]
        else:
            lines += [f"<img src='{_type.replace(' ', '_')}.png' style='width:38%'></img>\n"]
    f.writelines(lines)
