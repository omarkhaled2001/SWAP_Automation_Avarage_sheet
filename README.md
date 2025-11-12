# 📘 SWAP — Robot & Site Excel Generator

## 📖 Overview

This script automatically generates an Excel file summarizing **robots’ and sites’ performance data** based on input CSV file:

- `User Inputs.csv`

It creates a structured Excel sheet with calculated columns and formulas ready for data entry and reporting.

---

## ⚙️ Prerequisites

Before running the script, make sure your environment meets these requirements:

### 🧩 1. Installed Software

- **Python 3.8+**
- **Pip** (Python package manager)

### 📦 2. Required Python Libraries

Run the following command in your terminal or command prompt:

```bash
pip install openpyxl
```

### 📁 3. Input File

You must have the `User Inputs.csv` CSV file in the **same directory** as `generate.py` it has 3 columns:

#### `Column 1: Robots SN`

A list of robot serial numbers (one per line).
Example:

```
SN001
SN002
SN003
```

#### `Column 2: Sites Name`

A list of site names (one per line).
Example:

```
Site_A
Site_B
Site_C
```

#### `Column 3: Cumulative Acres since Date`

A String of Cumulative Acres since Date (one value @ the second line).
Example:

```
10/20/2025
```

> ⚠️ **Important:**
>
> - The first Row cells in each Column will be ignored by the script (so it can contain a header like “Robot SN” or “Site Name” or “Cumulative Acres since...?”).
> - Do **not** leave completely empty lines inside the file.

---

## 🚀 How to Use

1. Place all two files in the same folder:

   ```
   generate.py
   User Inputs.csv
   ```

2. Open a terminal or command prompt in that folder.

3. Run the script:

   ```bash
   python generate.py
   ```

4. The script will:

   - Read CSV file.
   - Validate that the total number of columns fits within Excel’s limit.
   - Create a structured Excel file named:

     ```
     N_Robots_M_Sites.xlsx
     ```

     (for example: `5_Robots_3_Sites.xlsx`)

5. If an older version of the same file exists, it will automatically be deleted before generating the new one.

---

## 🧮 Output Example

The generated Excel sheet will include:

- Robot-related columns: Acres, Time logs, Site names.
- Summary columns: Total acres, Average time, and per-site statistics.
- Automatic Excel formulas for cumulative and average values.
- Chart-ready data columns labeled as **“Chart’s Robots in Sites”**.

---

## 🚧 Boundary Conditions

To ensure Excel can handle the output:

```
7 * N + N * M + 6 * M ≤ 18275
```

Where:

- `N` = number of robots
- `M` = number of sites

If this boundary is exceeded, the script will show:

```
Boundaries was broken!
Can not generate final file according to large number of Robots or Site
please fit the Boundary equation: 7 * N + N * M + 6 * M =< 18275
```

---

## 🪲 Troubleshooting

| Problem                                           | Possible Cause          | Solution                                       |
| ------------------------------------------------- | ----------------------- | ---------------------------------------------- |
| `ModuleNotFoundError: No module named 'openpyxl'` | Library not installed   | Run `pip install openpyxl`                     |
| “Boundaries was broken!”                          | Too many robots/sites   | Reduce number of entries                       |
| Output file not created                           | Missing input CSV file  | Make sure CSV file exist in the same directory |
| Blank or missing headers                          | Empty first line in CSV | Remove extra empty lines                       |

---

## 🏁 Example

**Input:**

- `User Inputs.csv`

  ```
  Robot's Serial Number,Site's Name,Cumulative Acres since ... ?
  R1,Site_A,10/20/2025
  R2,Site_B,
  R3,,
  ```

**Run:**

```bash
python generate.py
```

**Output File:**
`3_Robots_2_Sites.xlsx`

---

## 📄 License

This script is free to use and modify for internal or educational purposes.
