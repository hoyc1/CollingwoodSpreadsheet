import pandas as pd
import numpy as np
import argparse
import xlrd
import xlwt

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

for year in np.unique(df["Year"]):
    bool = year == df["Year"]
    average = np.mean(
        [np.sum(df["Points value of sanction"][bool][df["Pupil Name"] == student].astype(np.float32)) for student in unique_names]
    )
    print(f"Average number of points per student in {year}: {average}")
