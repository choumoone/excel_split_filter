# Excel_split_filter
Excel_split_filter 是一个基于 Python 和 Tkinter 的桌面应用程序，用于处理 Excel 文件。它允许用户选择一个 Excel 文件进行拆分和筛选，并将结果保存到新的工作簿中。

## 功能
- 选择和加载 Excel 文件： 用户可以通过图形界面选择 Excel 文件（支持 .xlsx 和 .xls 格式），然后加载到程序中。
- 拆分数据： 程序会根据用户选择的列来拆分数据，每个唯一值生成一个新的数据帧。
- 筛选数据： 用户可以为拆分后的数据添加筛选条件，支持多个逻辑组合（AND/OR）和操作符（如等于、不等于、包含等）。
- 导出结果： 筛选和拆分后的数据将被保存到新的 Excel 工作簿中，每个筛选组的结果在单独的工作表上。
## 使用说明
- 启动程序后，点击“选择Excel文件”按钮，选择需要处理的Excel文件。
- 程序会提示选择一个列进行拆分。选择后，它将基于该列的唯一值拆分数据。
- 之后，程序会询问是否需要对数据进行筛选。如果选择“是”，则可以添加筛选条件。
- 每个筛选条件可以选择列、操作符和输入比较值。筛选条件可以组合使用“AND”和“OR”逻辑。
- 设置好筛选条件后，点击“应用筛选”，程序将处理数据并保存到新的Excel工作簿中。
## 环境要求
- Python
- Pandas
- Tkinter
- xlswriter（如果使用 xlsxwriter 作为 pandas ExcelWriter的引擎）

# Excel_split_filter
Excel_split_filter is a desktop application built with Python and Tkinter, designed for processing Excel files. It enables users to select an Excel file to split and filter, then saves the filtered data into a new workbook.

## Features
- Select and Load Excel Files: Users can choose an Excel file (supports .xlsx and .xls formats) through a graphical interface and load it into the program.
- Data Splitting: The program splits the data based on a column selected by the user, creating a new data frame for each unique value.
- Data Filtering: Users can apply filters to the split data, supporting multiple logical combinations (AND/OR) and operators (such as equals, not equal, contains, etc.).
- Export Results: The filtered and split data are saved into a new Excel workbook, with each set of filter results on separate sheets.
## Usage
- After starting the program, click the “Select Excel File” button to choose the Excel file you want to process.
- The program will prompt you to select a column for splitting. After selection, it will split the data based on the unique values of that column.
- Then, the program asks if you want to filter the data. If “Yes” is selected, you can add filter conditions.
- For each filter condition, you can select a column, an operator, and input a value for comparison. Filter conditions can be combined with “AND” and “OR” logic.
- Once the filter conditions are set, click “Apply Filters”, and the program will process the data and save it to a new Excel workbook.
## Requirements
- Python
- Pandas
- Tkinter
- xlsxwriter (if using xlsxwriter as the engine for pandas ExcelWriter)
