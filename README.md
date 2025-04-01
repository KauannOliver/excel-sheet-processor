# **Excel Sheet Processor - Automated Data Formatter**

This project was developed to automate the process of reading, organizing, and formatting Excel spreadsheet data, especially in business contexts where large amounts of financial or operational data are exported and need structured formatting. Built with Python and the Flet framework, the system provides a simple user interface to import, process, and export Excel files with properly grouped and styled data.

---

## **KEY FEATURES**

### 1. **Interactive Desktop Interface**
- Built with Flet, the app provides a modern and intuitive desktop UI.
- Allows users to select input and output Excel files using native file pickers.

### 2. **Automated Excel Formatting**
- Uses `pandas` to read and manipulate Excel data.
- Groups rows based on a concatenation of the first two columns.
- For each unique group, creates or updates a dedicated worksheet in the output file.

### 3. **Custom Sheet Creation and Safety**
- Automatically sanitizes worksheet names to ensure compatibility with Excel's 31-character limit.
- Removes invalid characters and limits the name length to prevent file corruption.

### 4. **Chunk Processing for Performance**
- Processes large Excel files in chunks of 1000 rows to reduce memory usage and improve performance.

### 5. **Data Validation and Safety Checks**
- Verifies if input and output files exist before proceeding.
- Handles missing values and avoids crashes due to incomplete or corrupted data.

### 6. **Styling and Number Formatting**
- Applies consistent number formats (currency, percentage, dates) across cells.
- Uses Openpyxl to style cells and apply a clean layout.

### 7. **Status Feedback to the User**
- Displays real-time status messages in the interface (e.g., "Loading", "Processing", "Success").
- Alerts for missing files or processing errors with color-coded messages.

---

## **TECHNOLOGIES USED**

### 1. **Python**
- The main programming language powering the logic and Excel processing.

### 2. **Flet**
- For building the user interface in a modern desktop style.

### 3. **Pandas**
- Handles data manipulation and grouping.

### 4. **Openpyxl**
- Used to read, write, format, and style Excel files.

### 5. **Regex (re)**
- Cleans and sanitizes worksheet names from invalid characters.

### 6. **OS & Time**
- For file checks and measuring execution time.

---

## **CONCLUSION**

The **Excel Sheet Processor** is a practical tool for anyone who needs to regularly organize and clean large Excel files. It automates grouping, styling, and sheet management, turning raw data exports into professionally formatted documents with minimal effort. With an easy-to-use UI and robust back-end logic, it's an ideal solution for analysts, accountants, and operations teams.

This tool eliminates repetitive manual formatting, saves time, and improves the consistency of your Excel reports.

