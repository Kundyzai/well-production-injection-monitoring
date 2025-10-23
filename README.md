# 🛠️ Well Production & Injection Monitoring Automation

## 🌍 Overview
This project automates the daily monitoring and data processing for oil production wells (CPU) and water injection wells (Injection), significantly reducing manual data entry time from 20–30 minutes to just 2–3 minutes per day.

---

## ⚙️ How It Works

### 1. CPU Monitoring (`cpu_monitor.py`)
- Watches the folder for new files named like `CPU_Production_EN DD.MM.YYYY.xlsx`.
- Extracts daily well data (oil, water, gas, RPM, etc.).
- Updates or creates a new date column in `CPU_update.xlsm`.
- Keeps a backup copy (`CPU_update_backup.xlsm`) before saving.
- Logs all actions in the console.

### 2. Injection Monitoring (`injection_monitor.py`)
- Monitors new injection reports (`Daily_Production_EN_CB DD.MM.YYYY.xlsx`).
- Reads well data (`well_92–well_101`) and transfers to `Injection.xlsx`.
- Automatically adds a new date row and copies Excel formulas for continuity.
- Updates columns for **Online Hours**, **THP**, and **Water Injection**.

---

## 🧩 Folder Structure

```
python_project/
│
├── BAT/
│   ├── start_cpu_monitor.bat
│   └── start_injection_monitor.bat
├── input/
│   ├── CPU_Production_EN 21.10.2025.xlsx
│   └── Daily_Production_EN_CB 30.09.2025.xlsx
├── macros/
│   ├── HighlightRPMChanges.vba
│   └── TransferData.vba
├── output/
│   ├── CPU_update.xlsm
│   └── Injection.xlsx
├── python_scripts/                   
│   ├── cpu_monitor.py
│   └── injection_monitor.py
├── README.md                          
└── requirements.txt                   
```
---

## 🚀 How to Use

### Step 1. Install Python (if not installed)
Download and install the latest [Python](https://www.python.org/downloads/) version (3.10+ recommended).  
During installation, check **“Add Python to PATH”**.

### Step 2. Install dependencies
In the project folder, run:

```bash
pip install -r requirements.txt
```

### Step 3. Configure folder paths
Inside each Python file (`cpu_monitor.py` and `injection_monitor.py`), update the following lines:

```python
PATH = r"C:\Your\Project\Path\CPU_update"
FOLDER_PATH = r"C:\Your\Project\Path\Injection folder"
```

Set them to your actual working directory.

### Step 4. Start automation
You can run either of these methods:

#### Option A – Manual run
Open a terminal (Command Prompt) and type:
```bash
python cpu_monitor.py
python injection_monitor.py
```

#### Option B – Use the BAT files
Just double-click:
- `start_cpu_monitor.bat`
- `start_injection_monitor.bat`

> 💡 You can also add these `.bat` files to **Windows Startup** so monitoring begins automatically after login.

### Step 5. Done!
Each time a new Excel report is placed in the folder, the script:
- Detects it automatically,
- Processes the data,
- Updates the target Excel file,
- Saves and backs up changes.

---
## 🧾 Excel Macros (Optional)

The project includes two VBA macros for enhanced data analysis:

1. **[HighlightRPMChanges.vba](./macros/HighlightRPMChanges.vba)** — Highlights RPM value changes between days in orange.  
   💬 *This allows you to quickly see if a well has stopped, resumed, or if its RPM has increased — helping to track operational performance.*  

2. **[TransferData.vba](./macros/TransferData.vba)** — Transfers production data between different sheets for analysis.  
   💬 *You can use this macro to move data from specific horizons or formations for separate analysis or reporting.*  

> 💡 You can also create **your own macros** for additional analysis using these as templates — for example, to track changes in the **Gas Factor (GOR)** or other production indicators over time.

---

## 🧠 Key Benefits

✅ Saves 20–30 minutes of manual work daily  
✅ Eliminates copy-paste errors  
✅ Automatically keeps backups  
✅ Easy to extend or customize  
✅ Runs silently in the background  

---

## 👩‍💻 Author
Project by **[Kundyzay]**  
Automation for production and injection well monitoring — saving time for real analysis.

💡 Customization Tip: By using my code templates — adjusting file paths, well names, data columns (oil, gas, water, etc.), and file names — you can easily automate the data processing for your own oilfield or project.
