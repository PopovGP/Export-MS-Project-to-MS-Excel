Attribute VB_Name = "ExportExcel"
'For the future: Option Explicit

Dim myDepth As Integer
Enum XlLineStyle
     xlContinuous = 1           ' Continuous line.
     xlDash = -4115             ' Dashed line.
     xlDashDot = 4              ' Alternating dashes and dots.
     xlDashDotDot = 5           ' Dash followed by two dots.
     xlDot = -4118              ' Dotted line.
     xlDouble = -4119           ' Double line.
     xlLineStyleNone = -4142    ' No line.
     xlSlantDashDot = 13        ' Slanted dashes.
 End Enum
 
Enum XlBordersIndex
    xlDiagonalDown = 5          ' Border running from the upper-left corner to the lower-right of each cell in the range.
    xlDiagonalUp = 6            ' Border running from the lower-left corner to the upper-right of each cell in the range.
    xlEdgeBottom = 9            ' Border at the bottom of the range.
    xlEdgeLeft = 7              ' Border at the left edge of the range.
    xlEdgeRight = 10            ' Border at the right edge of the range.
    xlEdgeTop = 8               ' Border at the top of the range.
    xlInsideHorizontal = 12     ' Horizontal borders for all cells in the range except borders on the outside of the range.
    xlInsideVertical = 11       ' Vertical borders for all the cells in the range except borders on the outside of the range.
End Enum

Enum XlHAlign
    xlHAlignCenter = -4108              ' Center.
    xlHAlignCenterAcrossSelection = 7   ' Center across selection.
    xlHAlignDistributed = -4117         ' Distribute.
    xlHAlignFill = 5                    ' Fill.
    xlHAlignGeneral = 1                 ' Align according to data type.
    xlHAlignJustify = -4130             ' Justify.
    xlHAlignLeft = -4131                ' Left.
    xlHAlignRight = -4152               ' Right.
End Enum

Enum XlVAlign
    xlVAlignBottom = -4107              ' Bottom
    xlVAlignCenter = -4108              ' Center
    xlVAlignDistributed = -4117         ' Distributed
    xlVAlignJustify = -4130             ' Justify
    xlVAlignTop = -4160                 ' Top
End Enum

Enum XlBorderWeight
    xlHairline = 1                      ' Hairline (thinnest border).
    xlMedium = -4138                    ' Medium.
    xlThick = 4                         ' Thick (widest border).
    xlThin = 2                          ' Thin.
End Enum

'for grouping
Sub myTree(myTask As Task)

    Dim myIdent As Integer
    
    If Not myTask Is Nothing Then
        With myTask
            If .Active Then
                
                myDepth = myDepth + 1
                
                If .OutlineChildren.Count > 0 Then 'we need to group
                    For myIdent = 1 To .OutlineChildren.Count
                        Call myTree(.OutlineChildren(myIdent))
                    Next myIdent
                End If
            End If
        End With
    End If
End Sub


Function myD(D As Variant)
    myD = DateSerial(Year(D), Month(D), Day(D))
End Function

Sub myBorders(mySelection As Variant, myLineStyle As Long, myWeight As Long)
    With mySelection.Borders(xlEdgeLeft)
        .LineStyle = myLineStyle
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = myWeight
    End With
    With mySelection.Borders(xlEdgeTop)
        .LineStyle = myLineStyle
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = myWeight
    End With
    With mySelection.Borders(xlEdgeBottom)
        .LineStyle = myLineStyle
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = myWeight
    End With
    With mySelection.Borders(xlEdgeRight)
        .LineStyle = myLineStyle
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = myWeight
    End With
    With mySelection.Borders(xlInsideVertical)
        .LineStyle = myLineStyle
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = myWeight
    End With
    With mySelection.Borders(xlInsideHorizontal)
        .LineStyle = myLineStyle
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = myWeight
    End With
End Sub
    
Sub myFormat(myApp As Variant, mySheet As Variant, myRange As Object, myEndRange As Object, myBold As Boolean, mySize As Integer, myColor As Long)
    
    With myRange
        .Font.Name = "Times New Roman"
        .Font.Bold = myBold
        .Font.Size = mySize
        .Interior.Color = myColor
    End With
    mySheet.Range(myRange, myEndRange).Select

     
    With myApp.Selection
        .HorizontalAlignment = XlHAlign.xlHAlignCenter
        .VerticalAlignment = XlVAlign.xlVAlignCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    
End Sub
    
Sub myMainFormat(myRange As Object, myBold As Boolean)
    With myRange
        .Font.Name = "Times New Roman"
        .Font.Bold = myBold
    End With
End Sub
    
Sub GanttFormat(myApp As Variant, mySheet As Variant, myCellStart As Object, myCellEnd As Object, myColor As Variant)

    mySheet.Range(myCellStart, myCellEnd).Select
    With myApp.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlDash
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = XlBorderWeight.xlThin
    End With
    With myApp.Selection.Borders(xlEdgeTop)
        .LineStyle = xlDash
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = XlBorderWeight.xlThin
    End With
    With myApp.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDash
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = XlBorderWeight.xlThin
    End With
    With myApp.Selection.Borders(xlEdgeRight)
        .LineStyle = xlDash
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = XlBorderWeight.xlThin
    End With
    
    
    myApp.Selection.Interior.Color = myColor
    
End Sub
   
Function AutoIndent(Level As Integer) As String
    Dim res As String
    Dim i As Integer
        
    res = ""
    
    For i = 1 To Level
        res = res + " "
    Next i
    AutoIndent = res
End Function
      
Sub ExportExcel()


    If ActiveProject.Tasks.Count = 0 Then
        MsgBox "Your project is empty!"
        Exit Sub
    End If
    
    Dim currentLine As Integer
    
    Dim myTask As Task
    
 

    
'late binding
' Dim Excel As Object
' Dim workbook As Object
'
'Set Excel = CreateObject("Excel.Application")
'set workbook = excel.
'Excel.Visible = True
'Excel.Close
    
    'using late binding
    Dim excelapp As Object      'for early binding: Dim excelapp As Excel.Application
    Dim workbook  As Object     'for early binding: Dim workbook As Excel.workbook
    Dim mySheet As Object       'for early binding: Dim mySheet As Excel.Worksheet
    
    
    Dim myBold As Boolean
    Dim ganttColumn As Integer 'Column where Gantt chart begins
    
    Set excelapp = CreateObject("Excel.Application") 'for early binding: New Excel.Application
    excelapp.ScreenUpdating = False
    Set workbook = excelapp.Workbooks.Add()
    Set mySheet = workbook.worksheets(1)
    mySheet.Name = "Gantt"
    ganttColumn = 9         ' <= tune this number if you want to load more columns from ms project. It is an offest column from wich Gantt chart starts
    
'for gantt chart
    'Dim myArr As Object
    Dim myStartDate As Date
    Dim myEndDate As Date
    Dim numberOfDays As Long
    Dim myArr() As Date
    
    
    myArr = StartAndEnd()
    myStartDate = myArr(0)
    myEndDate = myArr(1)
    numberOfDays = myEndDate - myStartDate
    excelapp.Visible = True

    With mySheet.Outline
        .AutomaticStyles = False
        .SummaryRow = xlAbove
        .SummaryColumn = xlRight
    End With
    
    
    'headers for tasks table
    mySheet.Cells(1, 1).Value = "¹"
    mySheet.Cells(1, 2).Value = Application.FieldConstantToFieldName(PjField.pjTaskUniqueID) '"Unique ID"
    mySheet.Cells(1, 3).Value = Application.FieldConstantToFieldName(PjField.pjTaskName) '"Name"
    mySheet.Cells(1, 4).Value = Application.FieldConstantToFieldName(PjField.pjTaskStartText) '"Start"
    mySheet.Cells(1, 5).Value = Application.FieldConstantToFieldName(PjField.pjTaskFinish)  '"Finish"
    mySheet.Cells(1, 6).Value = Application.FieldConstantToFieldName(PjField.pjTaskDurationText) '"Duration"
    mySheet.Cells(1, 7).Value = Application.FieldConstantToFieldName(PjField.pjTaskResourceNames) '"Resource names"
    mySheet.Cells(1, 8).Value = Application.FieldConstantToFieldName(PjField.pjTaskPredecessors)  '"Predecessors"
    mySheet.Cells(1, 9).Value = Application.FieldConstantToFieldName(PjField.pjTaskPercentComplete)  '"% complete"
    

    

    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 1), mySheet.Cells(4, 1), True, 11, RGB(223, 227, 232))
    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 2), mySheet.Cells(4, 2), True, 11, RGB(223, 227, 232))
    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 3), mySheet.Cells(4, 3), True, 11, RGB(223, 227, 232))
    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 4), mySheet.Cells(4, 4), True, 11, RGB(223, 227, 232))
    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 5), mySheet.Cells(4, 5), True, 11, RGB(223, 227, 232))
    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 6), mySheet.Cells(4, 6), True, 11, RGB(223, 227, 232))
    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 7), mySheet.Cells(4, 7), True, 11, RGB(223, 227, 232))
    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 8), mySheet.Cells(4, 8), True, 11, RGB(223, 227, 232))
    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 9), mySheet.Cells(4, 9), True, 11, RGB(223, 227, 232))
    
'some usefull freezing
    mySheet.Activate
    excelapp.ActiveWindow.FreezePanes = False
    mySheet.Cells(5, 4).Select
    excelapp.ActiveWindow.FreezePanes = True
    
    
'Gantt title =======================
    Dim myActiveTaskCount As Long
    Dim myDate As Date
    Dim myLastDate As Date
    Dim myGanttMonthCount As Long
    Dim myGanttYearCount As Long
    Dim myGanttWeekCount As Long
    Dim myLasti As Long
    Dim startMonth, startYear, startWeek As Integer
    Dim i As Long
    Dim FirstDayInWeek As Date
    Dim isWorkingDay As Boolean
    
    
    
    
    
    
    
'get number of active tasks
    myActiveTaskCount = 0
    For Each myTask In ActiveProject.Tasks
        If Not myTask Is Nothing Then
            If myTask.Active Then
                myActiveTaskCount = myActiveTaskCount + 1
            End If
        End If
    Next myTask
        
        
    myDate = myStartDate
    startMonth = Month(myDate)
    startYear = Year(myDate)
    startWeek = DatePart("ww", myDate, vbUseSystemDayOfWeek)
    
    myGanttMonthCount = 1
    myGanttYearCount = 1
    myGanttWeekCount = 1
    myLasti = 1
    myLastDate = myDate
        
    For i = 1 To numberOfDays + 1
        
        If Not startMonth = Month(myDate) Then
            
            mySheet.Range(mySheet.Cells(2, ganttColumn + myGanttMonthCount), mySheet.Cells(2, ganttColumn + i - 1)).Select
            With excelapp.Selection
                .HorizontalAlignment = XlHAlign.xlHAlignLeft
                .VerticalAlignment = XlVAlign.xlVAlignCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = True
                .Value = Format(myDate - 1, "mmmm yyyy")
                .Font.Name = "Times New Roman"
            End With
            Call myBorders(excelapp.Selection, xlContinuous, XlBorderWeight.xlThin)
            
            
            startMonth = Month(myDate)
            myGanttMonthCount = i
            
        End If
        
        If Not startYear = Year(myDate) Then
            
            
            mySheet.Range(mySheet.Cells(1, ganttColumn + myGanttYearCount), mySheet.Cells(1, ganttColumn + i - 1)).Select
            With excelapp.Selection
                .HorizontalAlignment = XlHAlign.xlHAlignLeft
                .VerticalAlignment = XlVAlign.xlVAlignCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = True
                .Value = Format(myDate - 1, "yyyy")
                .Font.Name = "Times New Roman"
            End With
            Call myBorders(excelapp.Selection, xlContinuous, XlBorderWeight.xlThin)
                        
            startYear = Year(myDate)
            myGanttYearCount = i
            
        End If
        
        If Not startWeek = DatePart("ww", myDate, vbUseSystemDayOfWeek) Then
            
            mySheet.Range(mySheet.Cells(3, ganttColumn + myGanttWeekCount), mySheet.Cells(3, ganttColumn + i - 1)).Select
            With excelapp.Selection
                .HorizontalAlignment = XlHAlign.xlHAlignLeft
                .VerticalAlignment = XlVAlign.xlVAlignCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = True
                
                If DatePart("ww", myDate - 1, vbUseSystemDayOfWeek) = 1 Then 'for 1st week only
                    FirstDayInWeek = DateSerial(Year(myDate), 1, 1)
                Else
'https://stackoverflow.com/questions/30026572/how-to-find-the-start-date-of-week-from-a-given-date
                    FirstDayInWeek = (myDate - 1) - Weekday(myDate - 1, vbUseSystemDayOfWeek) + 1
                    If myGanttWeekCount = 1 And FirstDayInWeek < myStartDate Then
                        FirstDayInWeek = myStartDate
                    End If
                End If
                .Value = Format(FirstDayInWeek, "Medium Date")
                
                .Font.Name = "Times New Roman"
                .Font.Size = 8
            End With
            Call myBorders(excelapp.Selection, xlContinuous, XlBorderWeight.xlThin)
            
            
            startWeek = DatePart("ww", myDate)
            myGanttWeekCount = i
        End If
        
        isWorkingDay = ActiveProject.Calendar.Period(myDate, myDate).Working
        
        If Not isWorkingDay Then
            mySheet.Range(mySheet.Cells(4, ganttColumn + i), mySheet.Cells(4 + myActiveTaskCount, ganttColumn + i)).Select
            excelapp.Selection.Interior.Color = RGB(223, 227, 232)
        End If
        
        mySheet.Cells(4, ganttColumn + i).Value = Left(WeekdayName(Weekday(myDate, vbUseSystemDayOfWeek), True, vbUseSystemDayOfWeek), 1) 'was myDate
        mySheet.Cells(4, ganttColumn + i).Select
        With excelapp.Selection
            .HorizontalAlignment = XlHAlign.xlHAlignCenter
            .VerticalAlignment = XlVAlign.xlVAlignCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .Font.Name = "Times New Roman"
            .Font.Size = 8
        End With
        
        
        myLasti = i
        myLastDate = myDate
        
        myDate = myDate + 1
 
        Next i
            
'close merged cells for gantt chart
'close month
        If Not myGanttMonthCount = myLasti + 1 Then
            mySheet.Range(mySheet.Cells(2, ganttColumn + myGanttMonthCount), mySheet.Cells(2, ganttColumn + myLasti)).Select
            With excelapp.Selection
                .HorizontalAlignment = XlHAlign.xlHAlignLeft
                .VerticalAlignment = XlVAlign.xlVAlignCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = True
                .Value = Format(myLastDate, "mmmm yyyy")
                .Font.Name = "Times New Roman"
            End With
            Call myBorders(excelapp.Selection, xlContinuous, XlBorderWeight.xlThin)
        End If
            
'close year
        If Not myGanttYearCount = myLasti + 1 Then
            mySheet.Range(mySheet.Cells(1, ganttColumn + myGanttYearCount), mySheet.Cells(1, ganttColumn + myLasti)).Select
            With excelapp.Selection
                .HorizontalAlignment = XlHAlign.xlHAlignLeft
                .VerticalAlignment = XlVAlign.xlVAlignCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = True
                .Value = Format(myLastDate, "yyyy")
                .Font.Name = "Times New Roman"
            End With
            Call myBorders(excelapp.Selection, xlContinuous, XlBorderWeight.xlThin)
        End If
            
'close week
        If Not myGanttWeekCount = myLasti + 1 Then
            mySheet.Range(mySheet.Cells(3, ganttColumn + myGanttWeekCount), mySheet.Cells(3, ganttColumn + myLasti)).Select
            With excelapp.Selection
                .HorizontalAlignment = XlHAlign.xlHAlignLeft
                .VerticalAlignment = XlVAlign.xlVAlignCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = True
'https://stackoverflow.com/questions/30026572/how-to-find-the-start-date-of-week-from-a-given-date
                .Value = Format(myLastDate - Weekday(myLastDate, vbUseSystem) + 1, "Medium Date")
                .Font.Name = "Times New Roman"
                .Font.Size = 8
            End With
            Call myBorders(excelapp.Selection, xlContinuous, XlBorderWeight.xlThin)
        End If
            
'end close
    
    currentLine = 4
    
    For Each myTask In ActiveProject.Tasks
        If Not myTask Is Nothing Then
            
            With myTask
                If .Active Then
                    
                   
                    currentLine = currentLine + 1
    
                    
                    Set excelrange = mySheet.Cells(currentLine, 1)
                    excelrange.Value = .ID 'myTask.UniqueID
                    
                    mySheet.Cells(currentLine, 2).Value = "'" + .OutlineNumber
                    mySheet.Cells(currentLine, 3).Value = AutoIndent(.OutlineLevel) + .Name
    
                    mySheet.Cells(currentLine, 4).Value = .StartText
   
                    mySheet.Cells(currentLine, 5).Value = .FinishText
                    
                    mySheet.Cells(currentLine, 6).Value = .DurationText
                    mySheet.Cells(currentLine, 7).Value = .ResourceNames
                    mySheet.Cells(currentLine, 8).Value = .Predecessors
                    mySheet.Cells(currentLine, 9).Value = .PercentComplete
                    
    
                    
                    
                    myBold = False
                    If .OutlineChildren.Count > 0 Then
                        myBold = True
                    End If
                    
                    Call myMainFormat(mySheet.Cells(currentLine, 1), myBold)
                    Call myMainFormat(mySheet.Cells(currentLine, 2), myBold)
                    Call myMainFormat(mySheet.Cells(currentLine, 3), myBold)
                    Call myMainFormat(mySheet.Cells(currentLine, 4), myBold)
                    Call myMainFormat(mySheet.Cells(currentLine, 5), myBold)
                    Call myMainFormat(mySheet.Cells(currentLine, 6), myBold)
                    Call myMainFormat(mySheet.Cells(currentLine, 7), myBold)
                    Call myMainFormat(mySheet.Cells(currentLine, 8), myBold)
                    Call myMainFormat(mySheet.Cells(currentLine, 9), myBold)
                    
    'start Doing Gantt chart
    '====================================================================
    'tasks
                    myColor = RGB(217, 225, 242)
                    dateOffset = myD(.Start) - myD(myStartDate)
                    myDuration = myD(.Finish) - myD(.Start) + 1
                    
                    If .Milestone Then
                        
                        mySheet.Cells(currentLine, ganttColumn + dateOffset + 1).Value = ChrW(&H25CA)
                        mySheet.Cells(currentLine, ganttColumn + dateOffset + 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        mySheet.Cells(currentLine, ganttColumn + dateOffset + 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                        
                    Else
                        If .OutlineLevel = 1 Then
    '  myColor = RGB(48, 84, 150)
    ' myColor = RGB(66, 96, 162)
                            myColor = RGB(78, 121, 198)
                        ElseIf .OutlineLevel = 2 Then
                            myColor = RGB(142, 169, 219)
                        ElseIf .OutlineLevel = 3 Then
                            myColor = RGB(180, 198, 231)
                        ElseIf .OutlineLevel = 4 Then
                            myColor = RGB(217, 225, 242)
                        Else
                            myColor = RGB(217, 225, 242)
                        End If
     
    'Dotted line
                       
                        Call GanttFormat(excelapp, mySheet, mySheet.Cells(currentLine, ganttColumn + dateOffset + 1), mySheet.Cells(currentLine, ganttColumn + dateOffset + myDuration), myColor)
                    End If
               
    'end doing Gantt chart ==============================================
                
                Else
                End If
            End With
        End If
        
    Next myTask 'loop through all tasks
                
    'Do some grouping =========================================================================================
    currentLine = 4
    For Each myTask In ActiveProject.Tasks
        
        If Not myTask Is Nothing Then
            With myTask
                If .Active Then
                    
                    currentLine = currentLine + 1
                    
                    If .OutlineLevel < 8 Then 'levels more than 9 are treated as usual lines due to MS Excel limitations
                        
                        If .OutlineChildren.Count > 0 Then 'we need to group
                            
                            myStartLine = currentLine
                            myDepth = 0
                            
                            For myIdent = 1 To .OutlineChildren.Count
                                Call myTree(.OutlineChildren(myIdent))
                                Next myIdent
                                
                                myEndLine = myStartLine + myDepth
                                
                                If Not myStartLine = myEndLine Then
                                    mySheet.Rows(Trim(Str(myStartLine + 1)) & ":" & Trim(Str(myEndLine))).Group
                                    
                                End If
                                
                            End If
                        End If
                        
                        
                    End If 'Maximum displayable outline count is reached or not reached
            End With
        End If
    Next myTask
    'end grouping =============================================================================================
                       
    '{begin do some formatting in columns where Gantt is located on a sheet ======================================
    For i = 1 To ganttColumn
        mySheet.Columns(i).AutoFit
        If mySheet.Columns(i).ColumnWidth > 100 Then 'if column is too wide and more than 100 points, set it's width to 100 points.
            mySheet.Columns(i).ColumnWidth = 100
        End If
    
    Next i
                            
    'Set to minimum width all columns where Gantt lines are
    For i = ganttColumn + 1 To numberOfDays + ganttColumn + 1
        mySheet.Columns(i).ColumnWidth = 1
    Next i
    '}end formatting Gantt columns ==================================================================================
    
    Call myResourceExport(excelapp, workbook, mySheet)
    
    Call myCalendarsExport(excelapp, workbook, mySheet)
    
    
    Set mySheet = Nothing
    
    workbook.worksheets(1).Activate
    excelapp.ScreenUpdating = True
    
                               
End Sub
                            
Sub ExportExcel_Monthly()


    If ActiveProject.Tasks.Count = 0 Then
        MsgBox "Your project is empty!"
        Exit Sub
    End If
    
    Dim currentLine As Integer
    
    Dim myTask As Task
    
 

    
'late binding
' Dim Excel As Object
' Dim workbook As Object
'
'Set Excel = CreateObject("Excel.Application")
'set workbook = excel.
'Excel.Visible = True
'Excel.Close
    
    'using late binding
    Dim excelapp As Object      'for early binding: Dim excelapp As Excel.Application
    Dim workbook  As Object     'for early binding: Dim workbook As Excel.workbook
    Dim mySheet As Object       'for early binding: Dim mySheet As Excel.Worksheet
    
    
    Dim myBold As Boolean
    Dim ganttColumn As Integer 'Column where Gantt chart begins
    
    Set excelapp = CreateObject("Excel.Application") 'for early binding: New Excel.Application
    excelapp.ScreenUpdating = False
    Set workbook = excelapp.Workbooks.Add()
    Set mySheet = workbook.worksheets(1)
    mySheet.Name = "Gantt"
    ganttColumn = 9         ' <= tune this number if you want to load more columns from ms project. It is an offest column from wich Gantt chart starts
    
'for gantt chart
    'Dim myArr As Object
    Dim myStartDate As Date
    Dim myEndDate As Date
    Dim numberOfDays As Long
    Dim myArr() As Date
    
    
    myArr = StartAndEnd()
    myStartDate = myArr(0)
    myEndDate = myArr(1)
    numberOfDays = myEndDate - myStartDate
    excelapp.Visible = True

    With mySheet.Outline
        .AutomaticStyles = False
        .SummaryRow = xlAbove
        .SummaryColumn = xlRight
    End With
    
    
    'headers for tasks table
    mySheet.Cells(1, 1).Value = "¹"
    mySheet.Cells(1, 2).Value = Application.FieldConstantToFieldName(PjField.pjTaskUniqueID) '"Unique ID"
    mySheet.Cells(1, 3).Value = Application.FieldConstantToFieldName(PjField.pjTaskName) '"Name"
    mySheet.Cells(1, 4).Value = Application.FieldConstantToFieldName(PjField.pjTaskStartText) '"Start"
    mySheet.Cells(1, 5).Value = Application.FieldConstantToFieldName(PjField.pjTaskFinish)  '"Finish"
    mySheet.Cells(1, 6).Value = Application.FieldConstantToFieldName(PjField.pjTaskDurationText) '"Duration"
    mySheet.Cells(1, 7).Value = Application.FieldConstantToFieldName(PjField.pjTaskResourceNames) '"Resource names"
    mySheet.Cells(1, 8).Value = Application.FieldConstantToFieldName(PjField.pjTaskPredecessors)  '"Predecessors"
    mySheet.Cells(1, 9).Value = Application.FieldConstantToFieldName(PjField.pjTaskPercentComplete)  '"% complete"
    

    

    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 1), mySheet.Cells(4, 1), True, 11, RGB(223, 227, 232))
    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 2), mySheet.Cells(4, 2), True, 11, RGB(223, 227, 232))
    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 3), mySheet.Cells(4, 3), True, 11, RGB(223, 227, 232))
    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 4), mySheet.Cells(4, 4), True, 11, RGB(223, 227, 232))
    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 5), mySheet.Cells(4, 5), True, 11, RGB(223, 227, 232))
    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 6), mySheet.Cells(4, 6), True, 11, RGB(223, 227, 232))
    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 7), mySheet.Cells(4, 7), True, 11, RGB(223, 227, 232))
    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 8), mySheet.Cells(4, 8), True, 11, RGB(223, 227, 232))
    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 9), mySheet.Cells(4, 9), True, 11, RGB(223, 227, 232))
    
'some usefull freezing
    mySheet.Activate
    excelapp.ActiveWindow.FreezePanes = False
    mySheet.Cells(5, 4).Select
    excelapp.ActiveWindow.FreezePanes = True
    
    
'Gantt title =======================
    Dim myActiveTaskCount As Long
    Dim myDate As Date
    Dim myLastDate As Date
    Dim myGanttMonthCount As Long
    'Dim myGanttYearCount As Long
    'Dim myGanttWeekCount As Long
    Dim myLasti As Long
    Dim startMonth As Integer
    Dim i As Long
    Dim FirstDayInWeek As Date
    Dim isWorkingDay As Boolean
    
    
    
    
    
    
    
    'get number of active tasks
    myActiveTaskCount = 0
    For Each myTask In ActiveProject.Tasks
        If Not myTask Is Nothing Then
            If myTask.Active Then
                myActiveTaskCount = myActiveTaskCount + 1
            End If
        End If
    Next myTask
        
        
    myDate = myStartDate
    startMonth = Month(myDate)
    'startYear = Year(myDate)
    'startWeek = DatePart("ww", myDate, vbUseSystemDayOfWeek)
    
    myGanttMonthCount = 1
    'myGanttYearCount = 1
    'myGanttWeekCount = 1
    myLasti = 1
    myLastDate = myDate
        
    For i = 1 To numberOfDays + 1
        
        If Not startMonth = Month(myDate) Then
            
            mySheet.Range(mySheet.Cells(1, ganttColumn + myGanttMonthCount), mySheet.Cells(4, ganttColumn + myGanttMonthCount)).Select
            With excelapp.Selection
                .HorizontalAlignment = XlHAlign.xlHAlignLeft
                .VerticalAlignment = XlVAlign.xlVAlignCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = True
                .Value = Format(myDate - 1, "mmmm yyyy")
                .Font.Name = "Times New Roman"
            End With
            Call myBorders(excelapp.Selection, xlContinuous, XlBorderWeight.xlThin)
            
            
            startMonth = Month(myDate)
            myGanttMonthCount = myGanttMonthCount + 1
            
        End If
        
       
        
        
        
        myLasti = i
        myLastDate = myDate
        
        myDate = myDate + 1
 
        Next i
'close month
    '    If Not myGanttMonthCount = Month(myLastDate) Then
            mySheet.Range(mySheet.Cells(1, ganttColumn + myGanttMonthCount), mySheet.Cells(4, ganttColumn + myGanttMonthCount)).Select
            With excelapp.Selection
                .HorizontalAlignment = XlHAlign.xlHAlignLeft
                .VerticalAlignment = XlVAlign.xlVAlignCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = True
                .Value = Format(myLastDate, "mmmm yyyy")
                .Font.Name = "Times New Roman"
            End With
            Call myBorders(excelapp.Selection, xlContinuous, XlBorderWeight.xlThin)
'        End If


            
'end close
    
    currentLine = 4
    
    For Each myTask In ActiveProject.Tasks
        If Not myTask Is Nothing Then
            
            With myTask
                If .Active Then
                    
                   
                    currentLine = currentLine + 1
    
                    
                    Set excelrange = mySheet.Cells(currentLine, 1)
                    excelrange.Value = .ID 'myTask.UniqueID
                    
                    mySheet.Cells(currentLine, 2).Value = "'" + .OutlineNumber
                    mySheet.Cells(currentLine, 3).Value = AutoIndent(.OutlineLevel) + .Name
    
                    mySheet.Cells(currentLine, 4).Value = .StartText
   
                    mySheet.Cells(currentLine, 5).Value = .FinishText
                    
                    mySheet.Cells(currentLine, 6).Value = .DurationText
                    mySheet.Cells(currentLine, 7).Value = .ResourceNames
                    mySheet.Cells(currentLine, 8).Value = .Predecessors
                    mySheet.Cells(currentLine, 9).Value = .PercentComplete
                    
    
                    
                    
                    myBold = False
                    If .OutlineChildren.Count > 0 Then
                        myBold = True
                    End If
                    
                    Call myMainFormat(mySheet.Cells(currentLine, 1), myBold)
                    Call myMainFormat(mySheet.Cells(currentLine, 2), myBold)
                    Call myMainFormat(mySheet.Cells(currentLine, 3), myBold)
                    Call myMainFormat(mySheet.Cells(currentLine, 4), myBold)
                    Call myMainFormat(mySheet.Cells(currentLine, 5), myBold)
                    Call myMainFormat(mySheet.Cells(currentLine, 6), myBold)
                    Call myMainFormat(mySheet.Cells(currentLine, 7), myBold)
                    Call myMainFormat(mySheet.Cells(currentLine, 8), myBold)
                    Call myMainFormat(mySheet.Cells(currentLine, 9), myBold)
                    
    'start doing Gantt chart
    '====================================================================
    'tasks
                    myColor = RGB(217, 225, 242)
                    'datediff
                    'DateDiff("m",  DateSerial(Year(myStartDate), Month(myStartDate), Day(1)), DateSerial(Year(.Start), Month(.Start), Day(1)))
                     
                    
                    monthOffsetStart = MonthDiff(myStartDate, .Start)
                    monthOffestEnd = MonthDiff(myStartDate, .Finish)
                   ' myDuration = myD(.Finish) - myD(.Start) + 1
                    
                    If .Milestone Then
                        
                        mySheet.Cells(currentLine, ganttColumn + monthOffsetStart + 1).Value = ChrW(&H25CA)
                        mySheet.Cells(currentLine, ganttColumn + monthOffsetStart + 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        mySheet.Cells(currentLine, ganttColumn + monthOffsetStart + 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                        
                    Else
                        If .OutlineLevel = 1 Then
    '  myColor = RGB(48, 84, 150)
    ' myColor = RGB(66, 96, 162)
                            myColor = RGB(78, 121, 198)
                        ElseIf .OutlineLevel = 2 Then
                            myColor = RGB(142, 169, 219)
                        ElseIf .OutlineLevel = 3 Then
                            myColor = RGB(180, 198, 231)
                        ElseIf .OutlineLevel = 4 Then
                            myColor = RGB(217, 225, 242)
                        Else
                            myColor = RGB(217, 225, 242)
                        End If
     
    'Dotted line
                       
                        Call GanttFormat(excelapp, mySheet, mySheet.Cells(currentLine, ganttColumn + monthOffsetStart + 1), mySheet.Cells(currentLine, ganttColumn + monthOffestEnd + 1), myColor)
                    End If
               
    'end doing Gantt chart ==============================================
                
                Else
                End If
            End With
        End If
        
    Next myTask 'loop through all tasks
                
    'Do some grouping =========================================================================================
    currentLine = 4
    For Each myTask In ActiveProject.Tasks
        
        If Not myTask Is Nothing Then
            With myTask
                If .Active Then
                    
                    currentLine = currentLine + 1
                    
                    If .OutlineLevel < 8 Then 'levels more than 9 are treated as usual lines due to MS Excel limitations
                        
                        If .OutlineChildren.Count > 0 Then 'we need to group
                            
                            myStartLine = currentLine
                            myDepth = 0
                            
                            For myIdent = 1 To .OutlineChildren.Count
                                Call myTree(.OutlineChildren(myIdent))
                                Next myIdent
                                
                                myEndLine = myStartLine + myDepth
                                
                                If Not myStartLine = myEndLine Then
                                    mySheet.Rows(Trim(Str(myStartLine + 1)) & ":" & Trim(Str(myEndLine))).Group
                                    
                                End If
                                
                            End If
                        End If
                        
                        
                    End If 'Maximum displayable outline count is reached or not reached
            End With
        End If
    Next myTask
    'end grouping =============================================================================================
                       
    '{begin do some formatting in columns where Gantt is located on a sheet ======================================
    For i = 1 To ganttColumn
        mySheet.Columns(i).AutoFit
        If mySheet.Columns(i).ColumnWidth > 100 Then 'if column is too wide and more than 100 points, set it's width to 100 points.
            mySheet.Columns(i).ColumnWidth = 100
        End If
    
    Next i
                            
    'Set to minimum width all columns where Gantt lines are
    For i = ganttColumn + 1 To numberOfDays + ganttColumn + 1
        mySheet.Columns(i).AutoFit 'ColumnWidth = 1
    Next i
    '}end formatting Gantt columns ==================================================================================
    
    Call myResourceExport(excelapp, workbook, mySheet)
    
    Call myCalendarsExport(excelapp, workbook, mySheet)
    
    
    Set mySheet = Nothing
    
    workbook.worksheets(1).Activate
    excelapp.ScreenUpdating = True
    
                               
End Sub
                          
'Export resources
Sub myResourceExport(myExcelApp As Object, myWorkbook As Object, myWorksheet As Variant)
    'Resources' export section         =====================================================================================================================
    'Dim mySheet As Object       'for early binding: Dim mySheet As Excel.Worksheet
    
    Dim myResourceFields As New Collection
    
    myResourceFields.Add Item:=PjField.pjResourceID
    myResourceFields.Add (PjField.pjResourceCode)
    myResourceFields.Add (PjField.pjResourceName)
    myResourceFields.Add (PjField.pjResourceInitials)
    myResourceFields.Add (PjField.pjResourceGroup)
    myResourceFields.Add (PjField.pjResourceBaseCalendar)
    myResourceFields.Add (PjField.pjResourceBookingType)
    myResourceFields.Add (PjField.pjResourceEMailAddress)
    myResourceFields.Add (PjField.pjResourceStandardRate)
    myResourceFields.Add (PjField.pjResourceOvertimeRate)
    myResourceFields.Add (PjField.pjResourcePeakUnits)
    myResourceFields.Add (PjField.pjResourceMaxUnits)
    myResourceFields.Add (PjField.pjResourceType)
    myResourceFields.Add (PjField.pjResourceCost)
    myResourceFields.Add (PjField.pjResourceCostPerUse)
    myResourceFields.Add (PjField.pjResourceOvertimeCost)
    myResourceFields.Add (PjField.pjResourceNotes)

    Dim myResourceSheet As Object
    Set myResourceSheet = myWorkbook.worksheets.Add(, myWorksheet)
    myResourceSheet.Name = "Resources"
    
        
    'Begin making table header{
    myCounter = 1
                
    For Each myField In myResourceFields
        myResourceSheet.Cells(1, myCounter).Value = Application.FieldConstantToFieldName(myField)
        
        Call myFormat(myExcelApp, myResourceSheet, myResourceSheet.Cells(1, myCounter), myResourceSheet.Cells(4, myCounter), True, 11, RGB(223, 227, 232))
                
        myCounter = myCounter + 1
    Next myField
    '}end making table header

    '{begin make table data
    currentLine = 5         ' beginning data
                
    For Each myResource In ActiveProject.Resources
        If Not myResource Is Nothing Then
    
            myCounter = 1
                      
            For Each myField In myResourceFields
                myResourceSheet.Cells(currentLine, myCounter).Value = myResource.GetField(myField)
                        
                myCounter = myCounter + 1
            Next myField
            
            'for future analysis myResourceSheet.Cells(currentLine, myCounter).Value = myResource.Calendar
        
         
            currentLine = currentLine + 1
            
            ' Using calendars
            ' myResource.GetField (PjField.pjResourceBaseCalendar)
            
        End If
    Next myResource
    '}end making table date
    
    Set myResourceSheet = Nothing

   
    '}end of resources export section   =====================================================================================================================

End Sub

'Export calendars
Sub myCalendarsExport(myExcelApp As Object, myWorkbook As Object, myWorksheet As Variant)
  '{Begin calendars export section   =====================================================================================================================
    Dim myCalendarSheet As Object
    Dim myCalendar As Calendar
    
    Set myCalendarSheet = myWorkbook.worksheets.Add(, myWorksheet)
    myCalendarSheet.Name = "Calendars"
    
    Set myCalendar = ActiveProject.Calendar
    
    currentLine = 5
    myCounter = 2
    
    myCalendarSheet.Cells(currentLine - 1, myCounter).Value = "Date"
    myCalendarSheet.Cells(currentLine - 2, myCounter + 1).Value = myCalendar.Name
    myCalendarSheet.Cells(currentLine - 1, myCounter + 1).Value = "Working?"
    
    myCalendarSheet.Cells(currentLine - 1, myCounter + 2).Value = "Shift1.Start"
    myCalendarSheet.Cells(currentLine - 1, myCounter + 3).Value = "Shift1.Finish"

    myCalendarSheet.Cells(currentLine - 1, myCounter + 4).Value = "Shift2.Start"
    myCalendarSheet.Cells(currentLine - 1, myCounter + 5).Value = "Shift2.Finish"

    myCalendarSheet.Cells(currentLine - 1, myCounter + 6).Value = "Shift3.Start"
    myCalendarSheet.Cells(currentLine - 1, myCounter + 7).Value = "Shift3.Finish"
    
    myCalendarSheet.Cells(currentLine - 1, myCounter + 8).Value = "Shift4.Start"
    myCalendarSheet.Cells(currentLine - 1, myCounter + 9).Value = "Shift4.Finish"

    myCalendarSheet.Cells(currentLine - 1, myCounter + 10).Value = "Shift5.Start"
    myCalendarSheet.Cells(currentLine - 1, myCounter + 11).Value = "Shift5.Finish"
     
    Days = ActiveProject.ProjectFinish - ActiveProject.ProjectStart
    myDate = ActiveProject.ProjectStart
    
    For i = 0 To Days
        
        myCalendarSheet.Cells(currentLine, myCounter).Value = myDate
        
        With myCalendar.Period(myDate)
            myCalendarSheet.Cells(currentLine, myCounter + 1).Value = .Working
            myCalendarSheet.Cells(currentLine, myCounter + 2).Value = .Shift1.Start
            myCalendarSheet.Cells(currentLine, myCounter + 3).Value = .Shift1.Finish
        
            myCalendarSheet.Cells(currentLine, myCounter + 4).Value = .Shift2.Start
            myCalendarSheet.Cells(currentLine, myCounter + 5).Value = .Shift2.Finish
        
            myCalendarSheet.Cells(currentLine, myCounter + 6).Value = .Shift3.Start
            myCalendarSheet.Cells(currentLine, myCounter + 7).Value = .Shift3.Finish
            
            myCalendarSheet.Cells(currentLine, myCounter + 8).Value = .Shift4.Start
            myCalendarSheet.Cells(currentLine, myCounter + 9).Value = .Shift4.Finish
        
            myCalendarSheet.Cells(currentLine, myCounter + 10).Value = .Shift5.Start
            myCalendarSheet.Cells(currentLine, myCounter + 11).Value = .Shift5.Finish
        
        
        End With
        
        currentLine = currentLine + 1
        myDate = myDate + 1
        
    Next i
    
    myCalendarSheet.Columns(myCounter).AutoFit
    
    
    
    Set myCalendarSheet = Nothing
    Set myCalendar = Nothing
    
    '}End calendars export section   =====================================================================================================================
  
End Sub


'for gantt chart find minimum and maximum project's start and end dates
Function StartAndEnd()
'for the future to try this: https://www.project-systems.co.nz/VBA-Sample-Code/VBASampleTimescaleScroll.html
    Dim myStartDate As Date
    Dim myEndDate As Date
    Dim currentLine As Long
    
    currentLine = 1
    
    myStartDate = 0
    myEndDate = 0
    
    
    For Each myTask In ActiveProject.Tasks
        If Not myTask Is Nothing Then
           
            
            If myTask.Active Then
    
                currentLine = currentLine + 1
                
                
                If myStartDate = 0 Then
                    myStartDate = myD(myTask.Start)
                End If
                If myTask.Start < myStartDate Then
                    myStartDate = myD(myTask.Start)
                End If
                
                If myTask.Finish > myEndDate Then
                    myEndDate = myD(myTask.Finish)
                End If
                
            End If
    
        End If
        
        
    Next myTask

    
    Dim res(0 To 1) As Date
    res(0) = myStartDate
    res(1) = myEndDate
    
    StartAndEnd = res
    
End Function

Function MonthDiff(myDate1 As Variant, myDate2 As Variant)
  MonthDiff = DateDiff("m", DateSerial(Year(myDate1), Month(myDate1), 1), DateSerial(Year(myDate2), Month(myDate2), 1))
End Function


