# Setup Instructions for ShipmentIDFileSearch

## 1. Clone the repository
```
git clone <repo-url>
cd ShipmentIdFileSearch
```

## 2. Create and activate a virtual environment (if not already present)
```
python -m venv .venv
# On Windows PowerShell:
.venv\Scripts\Activate.ps1
# On Windows CMD:
.venv\Scripts\activate.bat
```

## 3. Install dependencies
```
pip install -r requirements.txt
```

## 4. Configure settings
- Run the app once and use the Settings button to set your SQL Server, database, search directory, and file editor.

## 5. Run the application
```
python ShipmentIDFileSearch.py
```

---

## Building the EXE with PyInstaller

1. Install PyInstaller (if not already installed):
```
pip install pyinstaller
```

2. Build the EXE using the .spec file:
```
pyinstaller ShipmentIDFileSearch.spec
```

- The EXE will be in the `dist` folder.
- You may need to adjust the .spec file for data files or icons as needed.
