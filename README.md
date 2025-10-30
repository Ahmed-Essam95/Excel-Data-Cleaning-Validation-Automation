# Excel-Data-Cleaning-Validation-Automation
This Python script automates the data cleaning, filtering, and validation process for Excel files used in telecom billingو data analysis workflows. It uses Pandas and OpenPyXL to handle, clean, and format Excel sheets, ensuring data accuracy and consistency through automated rules and color-coded validations.




# ⚙️ Key Features

File Accessibility Check
Automatically verifies whether the required Excel file exists, with up to three retry attempts.

Data Cleaning with Pandas

Removes empty rows and columns

Filters records based on specific conditions (e.g., valid Account Number, active Status, valid Msisdn)

Excludes test and dummy records

Automated Sheet Distribution

The script splits the cleaned data into multiple categorized Excel sheets

Excel Formatting & Highlighting (via OpenPyXL)

Automatically highlights invalid or suspicious cells with color codes:

Name Validation Using Reference Lists

Compares entries against two reference sheets (males and females) in Names-Array.xlsx to highlight gender-based name matches.

Dynamic File Paths & User Interaction

Prompts the user through each step with interactive console messages and waits for input where manual review is required.


# 🧩 Technologies Used

Python 3.x
Pandas — For data cleaning and manipulation
OpenPyXL — For Excel formatting and workbook management
OS / Time / String — For file handling, delays, and string operations


# 🧠 How It Works

Reads the main Excel file (excel file name.xlsx) from the current directory
Cleans and filters the data
Generates categorized Excel outputs (Dummy, English, E-Gated, Untitled, Main)
Highlights invalid or missing values
Compares names with the Names-Array.xlsx reference file
Saves all processed workbooks automatically

# 🧰 Author Notes

This automation was built to reduce manual Excel review time and increase data validation accuracy in telecom and enterprise data environments, It provides step-by-step progress updates and interactive checkpoints for safe user control.

# ✍️ Author: Ahmed Essam
