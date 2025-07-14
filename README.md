# SensorReportGenerator

A Python-based GUI tool to automate the generation of Excel reports for sensor measurements.

## 📌 Description

**SensorReportGenerator** is designed to help engineers, quality assurance teams, or lab technicians automatically create structured Excel reports for multiple sensors using a predefined template. The tool reads measurement data from a CSV or Excel file and fills a copy of the template for each sensor, preserving formatting, layout, and formulas.

## 🛠️ Features

- GUI interface for easy step-by-step usage
- Supports `.xlsx` and `.csv` data input
- Automatically:
  - Copies the Excel template
  - Inserts measurement values (`0bar`, `20bar`, `40bar`)
  - Fills the serial number and insulation resistance
  - Applies formulas for deviation calculations
- Generates one worksheet per sensor
- Saves the final file as a single consolidated Excel report

## 📥 Input Requirements

1. **Excel Template**  
   Must include the following cells ready to be filled:
   - `B3`: Sensor serial number
   - `C14:C16`: Measured values
   - `B17`: Insulation resistance
   - `E14:E16`: Will be auto-filled with formulas like `=1/$C$10*D14`
   - `A15`, `A16`: Will be filled with pressure calculation formulas

2. **Data File (`.csv` or `.xlsx`)**  
   Must contain a header row with the following columns:
   - `S/N`
   - `0bar`
   - `20bar`
   - `40bar`
   - `I-Widerstand`

## ▶️ How to Use

1. Run the script:  
   `python SensorReportGenerator.py`

2. In the GUI:
   - Step 1: Load the Excel template
   - Step 2: Load the sensor data file
   - Step 3: Click **Generate Final Report** and choose a location to save the output

## 💾 Output

An `.xlsx` file containing:
- One worksheet per sensor, named with the `S/N`
- Data filled and formulas applied
- Original template sheet removed

## 🔧 Dependencies

- Python 3.x
- `pandas`
- `openpyxl`
- `tkinter` (included with most Python distributions)

Install requirements:
```bash
pip install pandas openpyxl
