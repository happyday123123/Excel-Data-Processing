## Project Documentation: Excel Data Processing

### Introduction
This project focuses on processing data from an Excel file using Python. The code performs various operations such as extracting information, calculating values, and generating reports based on the data in the Excel file.



### Requirements
- Python 3.x
- Libraries: xlrd, xlwt, pandas, openpyxl

### Code Overview
The Python code consists of several steps to process the Excel data:

1. Importing Required Libraries: The necessary libraries (xlrd, xlwt, pandas, openpyxl) are imported to support Excel file handling and data manipulation.

2. Opening the Excel File: The code uses the xlrd library to open the specified Excel file ("test.xlsx") and access the desired sheet ("工作表2").

3. Finding the Starting Position: The FindTheHead() function is implemented to determine the starting position of the data in the sheet. It iterates through the cells until it finds a non-empty cell and returns the row and column indices.

4. Extracting Data: The code populates the "guzi" dictionary with data from the Excel sheet. It iterates over the columns starting from the position found in FindTheHead() and retrieves the name and price values. The name is used as the key in the dictionary, and the price is rounded and stored as the value.

5. Retrieving Unique Names: The code retrieves all unique names from the sheet and stores them in the "cn" set.

6. Populating Associations: The code iterates over the columns and rows, populating the "ren" dictionary with names and associated people. It checks if the name exists in the "cn" set and if so, adds the associated person to the list in the dictionary.

7. Counting Occurrences: The code counts the occurrences of each name in the "cn" set and stores the counts in the "count" dictionary.

8. Generating Associations: The code creates a string representation of the associated people for each name in the "cn" set and stores them in the "pai" dictionary.

9. Calculating Total Values: The code calculates the total value ("shen") for each name in the "cn" set by multiplying the count of associated people with their corresponding price from the "guzi" dictionary.

10. Creating DataFrames: The code converts the dictionaries into Pandas DataFrames: "data1" for associations, "data2" for occurrence counts, and "data3" for total values. These DataFrames are then concatenated into a single DataFrame named "data".

11. Writing to Excel: The code opens the Excel file in write mode using openpyxl and creates a new sheet named "新表2". It writes the DataFrame "data" to the newly created sheet in the Excel file.

12. Column Width Adjustment: The code defines a class "CXlAutofit" with a method for automatically adjusting column widths in the Excel file. It instantiates the class and calls the "style_excel" method to adjust the column widths in the "新表2" sheet.

### Usage
1. Install the required libraries: xlrd, xlwt, pandas, openpyxl.
2. Ensure that the Excel file "leo_20221011.xlsx" is located at the specified path.
3. Execute the Python code.
4. The processed data will be written to a new sheet named "新表2" in the Excel file.
5. Column widths in the "新表2" sheet will be automatically adjusted for better readability.

### Conclusion
This project demonstrates how to process data from an Excel file using Python. It provides functionality for extracting information, calculating values, and generating reports based on the data in the Excel file. The code can be customized and extended to suit specific requirements for data processing and analysis in Excel.