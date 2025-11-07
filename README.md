# Export MS Project Gantt chart to MS Excel
MS Visual Basic module to export tasks, Gantt chart and resources list from MS Project to MS Excel.
Cute output. Instant use. No setup required. 100% MS Project VBA.

**_Features:_**
  - Versatile
  - Exports all commonly used fields
  - Automated nice Gantt chart in Excel with color and border emphasizing
  - Tolerant to blank fields
  - Eight-levels' row grouping (MS Excel maximum) with different fonts
  - Support unlimited levels of indented tasks in MS Project
  - Automatically indents subtasks and does grouping
  - Gantt chart is created at right side of the sheet

**_Using:_**
  1. Download and open __'Empty_MS_Project_with_macros.mpp'__
  2. Allow Visual Basic macros when asked
  3. Open project file you want to export
  4. While in your project file, choose the __'View'__ tab, click __'Macros'__
  4. Choose __'ExportExcel'__
  5. Click __'Run'__


**_Exports task's fields:_**
  1. №
  2. Unique task ID
  3. Task name
  4. Task start date
  5. Task end date
  6. Duration
  7. Resource names
  8. Task predecessors
  9. % complete

**_Exports those resource list fields:_**
  1. №
  2. ID
  3. Code
  4. Name
  5. Initials
  6. Group
  7. Base
  8. Calendar
  9. Booking Type
  10. Email
  11. Address
  12. Standard Rate
  13. Overtime Rate
  14. Peak
  15. Max Units
  16. Type
  17. Cost
  18. Cost Per Use
  19. Overtime Cost

**_Examples:_**

  _1.1. Project file:_
    ![Project screenshot](https://raw.githubusercontent.com/PopovGP/Export-MS-Project-to-MS-Excel/master/Samples_and_Images/Initial_project_example.png)
     
  _1.2. Excel file:_
    ![Excel screenshot](https://raw.githubusercontent.com/PopovGP/Export-MS-Project-to-MS-Excel/master/Samples_and_Images/Result_excel_example.png)

  _2.1. Project file:_
    ![Project screenshot](https://raw.githubusercontent.com/PopovGP/Export-MS-Project-to-MS-Excel/master/Samples_and_Images/Initial_resourcesheet_example.png)
     
  _2.2. Excel file:_
    ![Excel screenshot](https://raw.githubusercontent.com/PopovGP/Export-MS-Project-to-MS-Excel/master/Samples_and_Images/Export_resourcesheet_example.png)

  _3.1. Project file:_
    ![Project screenshot](https://raw.githubusercontent.com/PopovGP/Export-MS-Project-to-MS-Excel/master/Samples_and_Images/Initial_calendar_example.png)
     
  _3.2. Excel file:_
    ![Excel screenshot](https://raw.githubusercontent.com/PopovGP/Export-MS-Project-to-MS-Excel/master/Samples_and_Images/Export_calendar_example.png)

**_Notes:_**
  1. Created for MS Project 2019 Professional.
  2. If errors found - please write me or do your branch and correct.
  3. Supports all languages.
  4. Dates are formatted using system locale settings.


**_If you want to use it in another project:_**

Use "ExportExcel.bas"
  1. Launch MS Project
  2. On the __View__ tab, click __'Macros'__, and then click __'Visual Basic'__
  3. In __ProjectGlobal__ right-click in __'Modules'__
  4. Click __'Import file'__ and choose __'ExportExcel.bas'__
  5. 'ExportExcel' module should appear in modules


This module is provided 'as-is' and comes with no warranty.
If any error found or you have a comment - feel free to write.
