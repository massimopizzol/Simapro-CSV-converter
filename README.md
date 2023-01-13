# Simapro-CSV-converter
Python script that converts a life cycle inventory (LCI) from Excel into Simapro CSV format.


**What the script does:**

* convert multiple sheets at once (create a .csv file for each sheet in the Excel file)
* both database processes and all types of exchanges can be specified


**What it doesn't do:**

* specify sub-compartment of emission (e.g. "high-population")
* specify uncertainty
* add comments

(all these parameters can be specified directly in the .csv file though)


**Requirements:**

* python 2.7 or higher
* pandas
* [xlrd](https://pypi.python.org/pypi/xlrd#downloads) module [installed](https://packaging.python.org/installing/) (default if python was installed via [Anaconda] (https://docs.continuum.io/anaconda/pkg-docs))

**Procedure**

1. Prepare the life cycle inventory in Excel, save it in the same folder as the Python script
2. From your IDE or shell, navigate into the folder and run the code to generate the .csv files (see below)
3. From SimaPro, use Import>File and the following settings:
	* File format: "SimaPro CSV"
	* Object link method: "Try to link imported objects to existing objects first"
	* CSV format separator: "Tab"
	* Other options: "Replace existing processes..."



**Compiling LCIs in Excel and worked example:**

The Excel file attached includes two fictional LCIs.
Running the code returns the two files LCI1.csv and LCI2.csv to be imported into SimaPro.

* Cells A1:D6 are fixed, do not insert rows or columns there
* Each column is a process of the foreground system, and is matched by an identically named row and in the same order (see that E7:G9 in the example is a square matrix)
* Use exact LCI database process names under the foreground system
* Use "Raw", "Air", "Water", "Soil", "Waste", "Social", "Economic"  to indicate exchanges
* Use "Wastetotreatment" to indicate database processes of the waste treatment category

Of course one can create his own dataframe directly in python (maybe I'll upload a tutorial at some point) with the same structure of the one above...or change the spcsv.py code to reflect other structures.


**Running the code**

The recommended method is to use the Python3 function to convert a pandas dataframe object in simapro.csv file. Requires the pandas library.

Save the __spcsv.py__ file in your working directory and use the __to_spcsv__ function. For example:

```python
from spcsv import *
import pandas as pd

spcsv_ready = pd.read_excel('LCI_Example.xlsx', 'LCI1', index_col = False, header = None)

to_spcsv(spcsv_ready, 'LCI1.csv')
```

If you want to runt his from command line:


* Using python 3 and the spcsv.py module. The `spccsv_comm_line.py` file is an example: `python spcsv_comm_line.py <myfile_name.xlsx>`

* The old way (not really recommended). This was designed for python 2.7 long time ago. From shell, navigate into the folder and type `python LCAscript_v1.2.py <myfile_name.xls>` Note that [xlrd does not support .xlsx files anymore](https://stackoverflow.com/questions/65250207/pandas-cannot-open-an-excel-xlsx-file) so to make this work save yuor excel file in .xls format.