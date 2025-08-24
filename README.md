# PowerShell Excel to CSV Splitter

This script automates the process of taking a multi-tab Excel file and splitting each worksheet into its own separate CSV file. This is useful for data processing, analysis, and ingestion into other systems that prefer CSV format.

---

### Purpose

In many business environments, data is often stored in multi-tab Excel workbooks. This script provides a quick and efficient way to break down these complex files into individual CSVs, which are easier to parse and manipulate with other scripts or applications.

---

### Usage

The script defines a single function, `Split-Xlsx`, which takes two arguments: the **path to the Excel file** and the **directory where you want to save the CSV files**.

**Function Syntax:**

```powershell
Split-Xlsx -excelFileName <PathToExcelFile> -csvLoc <PathToOutputDirectory>
```

**Example:**

To split an Excel file named `What.xlsx` located in `C:\Temp` and save the resulting CSVs back to the same directory, you would call the function like this:

```powershell
Split-Xlsx -excelFileName "C:\Temp\What.xlsx" -csvLoc "C:\Temp\"
```

---

### Notes

* **Prerequisites:** This script requires **Microsoft Excel** to be installed on the system where it is run, as it uses the Excel COM object for its operations.
* **Automation:** This script is a great starting point for data preprocessing and can be easily integrated into a larger automation workflow.
* **Process Termination:** The script includes a `stop-process` command to ensure the Excel COM object is fully closed after the operation is complete, preventing any lingering processes.
