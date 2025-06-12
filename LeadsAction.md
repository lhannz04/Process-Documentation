### Step 1: Template Preparation
- Delete all files inside the folders below:

```
📁 Your Main Folder
├── 📁 CDR - APN & LV
├── 📁 Generate Report
├── 📁 Master Data
├── 📄 2023.06 - Leads Action Summary Report.xlsx
├── 📄 Dashboard Generator - Phase 16.5.xlsm
└── ⚙️ folder_cleanUp.bat
```

- Run the **folder_cleanUp.bat**.

### Step 2: Raw Data Extraction - Macro

1. Open the **Dashboard Generator - Phase 16.5.xlsm**

2. Change the [Start Date] and [End Date] to Report Date.

3. Run [Clean Data] to remove the previous day data.

4. RUN Call Detail Report Fetcher.
    - Select 📂CDR - APN & LV
        - Select 📄Call Detail Report.txt file.
        - When confirmation message appears click [Ok]

5. Run [Masterlist Fetcher].
    - Select 📂Master Data
        - Select 📃Unified Master Data_Initial-LASG.xlsx

6. Run [Skip App & Others Consolidator].
    - Select 📂Generation Report.
        - Select 📃GENERATION_REPORT_DETAIL.xlsx

7. Run [Generated Loan Lock Report].
    - Select 📂Generation Report.
        - Select 📃Searched and Generated Leads.xlsb
        - Close Generator.xlsm and Save.


### Step 3: Final Data Modeling and Publish

1. Open [Leads Action Summary Report 1.2p.xlsm].
    - Go to worksheets [Leads Action RAW Template]

2.  Select 📂Generation Report.
    - Open 📃Searched and Generated Leads.xlsb (without RAW in the file name)
    - Kill the formula in Rows 1.
    - Select all data, entire row copy.

3. Go to [Leads Action Summary Report 1.2p.xlsm] worksheets [Leads Action RAW Template].
    - Select the Black rows entire row.
    - CTRL + '+' to insert the data.
    - Copy formatting.
    - Select the last row to verify that data is attach correctly.
    - Refresh All to refresh pivot tables.
    - Go to worksheets [Leads Action Summary Report]
    - Select the Report Date.
    - Close 📃Searched and Generated Leads.xlsb (without RAW in the file name)
    - Do not save.

4. Select [Leads Action Summary Report 1.2p.xlsm] worksheets [Leads Action RAW Template].
    - Run [Publish L.A.S. Report]
    - Ignore the error. Click [End].
    - Select the Report date.
    - Close and Save
