# Export-MS-Project-to-MS-Excel
MS Visual Basic module to export tasks, Gantt chart and resources list from MS Project to MS Excel.
Instant use. No setup required. 100% MS Project VBA.

**_Features:_**
  - Versatile
  - Exports all commonly used fields
  - Automated nice Gantt-chart in Excel with color and border emphasizing
  - Tolerant to blank fields
  - Eight-level's row grouping (MS Excel maximum) with different fonts
  - Support unlimited levels of indented tasks in MS Project
  - Automatically indents subtasks and do grouping
  - Gantt chart is created at right site of the sheet

**_Using:_**
  1. Download and open __'Empty_MS_Project_with_macros.mpp'__
  2. Allow Visual Basic macros when asked
  3. Open project file you want to import
  4. While in your project file, choose the __View__ tab, click __Macros__
  4. Choose __'ExportExcel'__
  5. Click __'Run'__


**_Exports task filelds:_**
  1. №
  2. Unique task ID
  3. Task name
  4. Task start date
  5. Task end datae
  6. Duration
  7. Resource names
  8. Task predecessors
  9. % complete

**_Export resources_**

**_Example:_**

  _Project file:_
    ![Project screenshot](https://raw.githubusercontent.com/PopovGP/Export-MS-Project-to-MS-Excel/master/Samples_and_Images/Initial_project_example.png)
     
  _Excel file:_
    ![Excel screenshot](https://raw.githubusercontent.com/PopovGP/Export-MS-Project-to-MS-Excel/master/Samples_and_Images/Result_excel_example.png)

**_Notes:_**
  1. Created for MS Project 2019 Professional.
  2. If errors found - please write me or do your branch and correct.
  3. Support all languages.
  4. Dates are formatted using system locale settings.


**_If you want to use it in another project:_**

Use "ExportExcel.bas"
  1. Launch MS Project
  2. On the __View__ tab, click __Macros__, and then click __Visual Basic__
  3. In __ProjectGlobal__ right-click in __Modules__
  4. Click __'Import file'__ and choose __'ExportExcel.bas'__
  5. 'ExportExcel' module should appear in modules


This module is provided 'as-is' and comes with no warranty.
If any error or comment - feel free to write.
