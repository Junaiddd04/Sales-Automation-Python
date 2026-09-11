# Automated Sales Reporting Pipeline
An automated ETL (Extract, Transform, Load) and reporting script that eliminates manual Excel tasks. This project reads messy daily sales CSVs, cleans the data, calculates key metrics, builds a fully formatted Excel dashboard with charts, archives the raw data, and emails the final report to stakeholders.

## 🚀 Key Features & Workflow
This script operates in a 6-step automated workflow:

**Data Ingestion & Logging:**

Initializes a background logger (Sales_Automation.log) to track every step of the process.

Scans the Incoming_Data directory and loads the pending CSV file into a Pandas DataFrame.

**Data Cleaning:**

Standardizes column names (lowercase with underscores) and sorts records by date.

Intelligently handles missing values by filling empty unit_price fields with the product's mean price, and empty quantity fields with the median.

Drops duplicate records to ensure data integrity.

**Feature Engineering & Aggregation:**

Calculates a new sales column (quantity * unit_price).

Extracts day_name and month_name from the date column.

Generates four distinct summary dataframes grouped by: Product, Region, Day, and Month.

**Excel Dashboard Generation:**

Uses openpyxl to write the raw data and summaries into separate worksheets.

Applies professional styling: Corporate blue headers (1F4E78) with bold white text, frozen top rows, auto-filters, and auto-adjusted column widths.

**Data Visualization:**

Injects conditional formatting (DataBars) into the summary sheets to highlight high/low values.

Programmatically generates and inserts Bar Charts (for products and months) and Line Charts (for regions and days) directly into the Excel grid.

**Archiving & Distribution:**

Saves the final workbook to an Output folder.

Moves the raw CSV file to an Archive_Data folder using shutil to prevent reprocessing.

Securely emails the finished Excel report to designated stakeholders using yagmail.

## 🛠️ Tech Stack
Python 3.x

**Pandas:** Data manipulation, cleaning, and aggregation.

**Openpyxl:** Excel workbook creation, formatting, and chart generation.

**Yagmail:** Simplified email distribution via SMTP.

**Standard Libraries:** os, shutil, pathlib, logging.


## 📂 Project Structure


├── Incoming_Data/             # Drop your raw CSV files here

├── Archive_Data/              # Processed CSVs are automatically moved here

├── Output/                    # Final Generated Excel reports are saved here

├── Sales_Automation.py        # The main execution script

├── Sales_Automation.log       # Auto-generated log file tracking pipeline health

└── README.md


## ⚙️ Setup & Configuration

**1. Install Dependencies**

pip install pandas openpyxl yagmail

**2. Configure Email Credentials**

Before running the script, you must configure the send_mail() function at the bottom of Sales_Automation.py:

Open the script and locate the send_mail() function.

Update the sender and receiver email addresses.

Update the password variable. Important: If using Gmail, you cannot use your standard login password. You must generate a 16-letter App Password from your Google Account Security settings.

**3. Run the Pipeline**

Place a raw sales CSV file into the Incoming_Data folder and run the script


## ⏱️ Daily Automation (Windows Task Scheduler)
This script is designed to be a fully hands-off pipeline. It is configured to run automatically every day using Windows Task Scheduler:

Open Task Scheduler in Windows and click Create Basic Task.

Name it "Daily Sales Automation" and set the Trigger to Daily (e.g., at 9:00 AM).

Set the Action to Start a program.

In the Program/script box, browse to your Python executable (e.g., C:\Python312\python.exe).

In the Add arguments box, type the name of the script: Sales_Automation.py.

In the Start in box, paste the absolute path to your project folder (e.g., C:\Users\YourName\Documents\Sales_Automation\).

Save the task.

As long as new data is dropped into the Incoming_Data folder, the pipeline will wake up daily, process the files, send the email, and go back to sleep.

**Author: Mohd Junaid.**
