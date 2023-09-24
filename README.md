# CollingwoodSpreadsheet
Repository to store code for reading and writing spreadsheets for Collingwood

## `script.py`

`script.py` is a lightweight script that reads and simplifies spreadsheets for Collingwood. It requires
`numpy`, `pandas`, `argparse`, `xlwt`, `xlrd`, `python>=3.7`. The easiest way of installing these dependencies
is via [miniconda](https://docs.conda.io/projects/conda/en/latest/user-guide/install/windows.html). This
script currently writes a new excel document containing a simplified table (by default called output.xls),
as well as a handy html file viewable from a browser (by default called output.html). The script can be
run with the following command from the terminal:

```bash
$ python script.py --input Rewards and Conduct Summary Report for Charlie.xls
```
