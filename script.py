import pandas as pd
import numpy as np
import argparse

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

df = pd.read_excel(opts.input)
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
    new_df["Total Value"][num] = np.sum(df["Points value of sanction"][bool])

new_df.to_excel(opts.output_excel, index=False)
new_df.to_html(buf=opts.output_html, index=False)

for year in np.unique(df["Year"]):
    bool = year == df["Year"]
    average = np.mean(
        [np.sum(df["Points value of sanction"][bool][df["Pupil Name"] == student]) for student in unique_names]
    )
    print(f"Average number of points per student in {year}: {average}")
