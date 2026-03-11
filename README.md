# Scrap & Inventory Loss Reconciliation Tool

This automation tool is designed to compare and reconcile "scrap" or "inventory loss" (zayi çıkış) records from two different database systems. Built with a user-friendly Tkinter graphical interface, it streamlines the manual comparison process by analyzing raw data exports and generating a formatted, easy-to-read Excel matrix.

## Features
* **Automated Data Processing:** Cleans and standardizes raw scrap/loss data from multiple sources.
* **Pivot & Matrix Generation:** Automatically creates pivot tables for both datasets and merges them into a single comparison matrix.
* **Visual Reporting:** Exports the final results to Excel with automated color-coding (blue and grey fills) to highlight data origins and discrepancies.
* **User-Friendly GUI:** A simple desktop application interface that allows users to process complex data comparisons with a few clicks.

## Technologies Used
* **Python**
* **Pandas:** Data manipulation, cleaning, and aggregation.
* **Openpyxl:** Automated Excel formatting and cell styling.
* **Tkinter:** Graphical user interface (GUI) development.

## How to Run
1. Clone this repository to your local machine.
2. Install the required dependencies:
   ```bash
   pip install pandas openpyxl
3. Run the application:
   ```bash
   python Zayi.py
4.Use the interface to upload your Excel files and generate the reconciliation matrix.
   
