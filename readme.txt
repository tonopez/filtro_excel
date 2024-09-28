# Help for Excel File Filter Application - Version 1.1.0

## Description
This application allows you to load an Excel file, dynamically filter its data, and now also edit all fields directly from the graphical user interface (GUI) based on `tkinter`.

## New in version 1.1.0
- All fields are now editable with a double-click.
- Improved handling of special column names.
- Increased overall stability and performance of the application.

## Instructions for Use

### 1. Load an Excel File
- Click on the **"Load Excel File"** button.
- Select the Excel file you want to load from your file system.

### 2. Filter Data
- Once the file is loaded, dropdown lists will be displayed for each column of the Excel file.
- Select the desired values in the dropdown lists to filter the data.
- The data will be filtered dynamically as you select the values.

### 3. View and Edit Results
- The filtered results will be displayed in the text box on the right.
- If there is only one result, the complete details of the record will be shown.
- If there are multiple results, the number of results found will be indicated.
- To edit any field, double-click on it in the results text box.
- An editing window will open where you can modify the field value.

### 4. Clear Fields
- Click on the **"Clear Fields"** button to reset all dropdown lists and display all data again.

### 5. Export Results
- Click on the **"Export Complete File"** button to save all data to a new Excel file.
- Click on the **"Export Filtered Results"** button to save only the filtered results to a new Excel file.
- Select the location and name of the file to save the results.

## Additional Features

### Dynamic Filter Update
- The options in the dropdown lists are dynamically updated based on the filters selected in other dropdown lists.

### Scroll Bar
- If the filtered data is extensive, you can use the horizontal and vertical scroll bars to navigate through the text box.

## Notes
- Make sure the Excel file you want to load is in `.xlsx` format.
- The application replaces `NaN` values in the Excel file with `"---"` and converts float values to integers if possible.
- All fields are now editable, including those with special column names or special characters.

---

This help file provides a quick guide on how to use the application and its main features. For any additional questions or suggestions, please contact the development team.
