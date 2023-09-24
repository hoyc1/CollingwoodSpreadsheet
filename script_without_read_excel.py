import pandas as pd
import numpy as np
import argparse
import xlrd

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

book = xlrd.open_workbook(opts.input)
sh = book.sheet_by_index(0)
data = []
for i in range(sh.nrows):
    row = [
        value for value, typ in zip(sh.row_values(i), sh.row_types(i))
    ]
    data.append(row)

data = np.array(data)
columns = data[0]
new_data = {}
for num, col in enumerate(columns):
    new_data[col] = data[:,num][1:]

df = pd.DataFrame.from_dict(new_data)
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

new_df.to_html(buf=opts.output_html, index=False)
