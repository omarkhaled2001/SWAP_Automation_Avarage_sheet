# 📘 SWAP — Robot & Site Excel Generator

## 📖 Overview

This script automatically generates an Excel file summarizing **robots’ and sites’ performance data** based on two input CSV files:

- `Robots SN.csv`
- `Sites Name.csv`

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

### 📁 3. Input Files

You must have the following CSV files in the **same directory** as `generate.py`:

#### `Robots SN.csv`

A list of robot serial numbers (one per line).
Example:

```
SN001
SN002
SN003
```

#### `Sites Name.csv`

A list of site names (one per line).
Example:

```
Site_A
Site_B
Site_C
```

> ⚠️ **Important:**
>
> - The first cell in each CSV will be ignored by the script (so it can contain a header like “Robot SN” or “Site Name”).
> - Do **not** leave completely empty lines inside the file.

---

## 🚀 How to Use

1. Place all three files in the same folder:

   ```
   generate.py
   Robots SN.csv
   Sites Name.csv
   ```

2. Open a terminal or command prompt in that folder.

3. Run the script:

   ```bash
   python generate.py
   ```

4. The script will:

   - Read both CSV files.
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

| Problem                                           | Possible Cause          | Solution                                             |
| ------------------------------------------------- | ----------------------- | ---------------------------------------------------- |
| `ModuleNotFoundError: No module named 'openpyxl'` | Library not installed   | Run `pip install openpyxl`                           |
| “Boundaries was broken!”                          | Too many robots/sites   | Reduce number of entries                             |
| Output file not created                           | Missing input CSV files | Make sure both CSV files exist in the same directory |
| Blank or missing headers                          | Empty first line in CSV | Remove extra empty lines                             |

---

## 🏁 Example

**Input:**

- `Robots SN.csv`

  ```
  Robot SN
  R1
  R2
  ```

- `Sites Name.csv`

  ```
  Site Name
  Site_A
  Site_B
  ```

**Run:**

```bash
python generate.py
```

**Output File:**
`2_Robots_2_Sites.xlsx`

---

## 📄 License

This script is free to use and modify for internal or educational purposes.
