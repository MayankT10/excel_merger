# Excel Merger

A simple Python script to merge multiple Excel files from a folder into a single Excel workbook.
(Each file will be added as a separate sheet.)

## Requirements

- Python 3.x
- pandas
- openpyxl

## Video tutorial of how to run 

![tutorial.gif](tutorial.gif)



## Steps to Run

1. Install dependencies:

    ```
    pip install pandas openpyxl
    ```

2. Clone this repo into your local folder:

    ```
    git clone https://github.com/MayankT10/excel_merger
    ```

3. Go to the `excel_merger` directory and run the script:

    ```
    cd excel_merger
    python main.py
    ```

4. Enter the folder path containing the Excel files when prompted.

5. Enter the desired output file name (with `.xlsx` extension).

The merged Excel file will be created in the current directory.