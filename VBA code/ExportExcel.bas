Attribute VB_Name = "ExportExcel"
Dim myDepth As Integer
Enum XlLineStyle
     xlContinuous = 1      'Continuous line.
     xlDash = -4115   'Dashed line.
     xlDashDot = 4    'Alternating dashes and dots.
     xlDashDotDot = 5     'Dash followed by two dots.
     xlDot = -4118    'Dotted line.
     xlDouble = -4119     'Double line.
     xlLineStyleNone = -4142  'No line.
     xlSlantDashDot = 13  'Slanted dashes.
 End Enum
 
Enum XlBordersIndex
    xlDiagonalDown = 5  ' Border running from the upper-left corner to the lower-right of each cell in the range.
    xlDiagonalUp = 6   ' Border running from the lower-left corner to the upper-right of each cell in the range.
    xlEdgeBottom = 9   ' Border at the bottom of the range.
    xlEdgeLeft = 7   'Border at the left edge of the range.
    xlEdgeRight = 10 ' Border at the right edge of the range.
    xlEdgeTop = 8   ' Border at the top of the range.
    xlInsideHorizontal = 12  'Horizontal borders for all cells in the range except borders on the outside of the range.
    xlInsideVertical = 11    'Vertical borders for all the cells in the range except borders on the outside of the range.
End Enum

Enum XlHAlign
    xlHAlignCenter = -4108   'Center.
    xlHAlignCenterAcrossSelection = 7    'Center across selection.
    xlHAlignDistributed = -4117  'Distribute.
    xlHAlignFill = 5     'Fill.
    xlHAlignGeneral = 1  'Align according to data type.
    xlHAlignJustify = -4130  'Justify.
    xlHAlignLeft = -4131     'Left.
    xlHAlignRight = -4152   ' Right.
End Enum

Enum XlVAlign
    xlVAlignBottom = -4107   'Bottom
    xlVAlignCenter = -4108   'Center
    xlVAlignDistributed = -4117  'Distributed
    xlVAlignJustify = -4130  'Justify
    xlVAlignTop = -4160  'Top
End Enum

Enum XlBorderWeight
    xlHairline = 1   'Hairline (thinnest border).
    xlMedium = -4138    'Medium.
    xlThick = 4  'Thick (widest border).
    xlThin = 2   'Thin.
End Enum

'for grouping
Sub myTree(myTask As Task)
    If Not myTask Is Nothing Then
        If myTask.Active Then
            
            myDepth = myDepth + 1
            
            If myTask.OutlineChildren.Count > 0 Then 'we need to group
                
                For myIdent = 1 To myTask.OutlineChildren.Count
                    Call myTree(myTask.OutlineChildren(myIdent))
                Next myIdent
            End If
        End If
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
    
    myRange.Font.Name = "Times New Roman"
    myRange.Font.Bold = myBold
    myRange.Font.Size = mySize
    myRange.Interior.Color = myColor
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
    myRange.Font.Name = "Times New Roman"
    myRange.Font.Bold = myBold
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
    res = ""
    For i = 1 To Level
        res = res + " "
    Next i
    AutoIndent = res
End Function
      
Sub ExportExcel()


    If ActiveProject.Tasks.Count = 0 Then
        MsgBox "Your project is emty!"
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
    Dim excelapp As Object  'for early binding: Dim excelapp As Excel.Application
    Dim workbook  As Object 'for early binding: Dim workbook As Excel.workbook
    Dim mySheet As Object     'for early binding: Dim mySheet As Excel.Worksheet
    
    
    Dim myBold As Boolean
    Dim ganttColumn As Integer 'Column where Gantt chart begins
    
    Set excelapp = CreateObject("Excel.Application") 'for early binding: New Excel.Application
    excelapp.ScreenUpdating = False
    Set workbook = excelapp.Workbooks.Add()
    Set mySheet = workbook.Worksheets(1)
    ganttColumn = 9         ' <= tune this number if you want to load more columns from ms project. It is an offest column from wich Gantt chart starts
    
'for gantt chart
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
    mySheet.Cells(1, 1).Value = "�"
    mySheet.Cells(1, 2).Value = Application.FieldConstantToFieldName(PjField.pjTaskUniqueID) '"Unique ID"
    mySheet.Cells(1, 3).Value = Application.FieldConstantToFieldName(PjField.pjTaskName) '"Name"
    mySheet.Cells(1, 4).Value = Application.FieldConstantToFieldName(PjField.pjTaskStartText) '"Start"
    mySheet.Cells(1, 5).Value = Application.FieldConstantToFieldName(PjField.pjTaskFinish)  '"Finish"
    mySheet.Cells(1, 6).Value = Application.FieldConstantToFieldName(PjField.pjTaskDurationText) '"Duration"
    mySheet.Cells(1, 7).Value = Application.FieldConstantToFieldName(PjField.pjTaskResourceNames) '"Resource names"
    mySheet.Cells(1, 8).Value = Application.FieldConstantToFieldName(PjField.pjTaskPredecessors)  '"Predecessors"
    mySheet.Cells(1, 9).Value = Application.FieldConstantToFieldName(PjField.pjTaskPercentComplete)  '"% complete"
    
    'for those who want to extent this code and don't know what is the correct name of a column
    'uncomment these lines and see the debug info
    'Dim myDebugArr(1337) As Long
    'DebugEnum (myDebugArr)
    'For i = 0 To 1337
    '    Debug.Print myDebugArr(i)
    'Next i
    'end uncommenting
    
    

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
'number of active tasks
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
            
            
            If myTask.Active Then
                
               
                currentLine = currentLine + 1

                
                Set excelrange = mySheet.Cells(currentLine, 1)
                excelrange.Value = myTask.ID 'myTask.UniqueID
                
                mySheet.Cells(currentLine, 2).Value = "'" + myTask.OutlineNumber
                mySheet.Cells(currentLine, 3).Value = AutoIndent(myTask.OutlineLevel) + myTask.Name
'  myTask.Start
                mySheet.Cells(currentLine, 4).Value = myTask.StartText
'  myTask.Finish
                mySheet.Cells(currentLine, 5).Value = myTask.FinishText
                
                mySheet.Cells(currentLine, 6).Value = myTask.DurationText
                mySheet.Cells(currentLine, 7).Value = myTask.ResourceNames
                mySheet.Cells(currentLine, 8).Value = myTask.Predecessors
                mySheet.Cells(currentLine, 9).Value = myTask.PercentComplete
                

                
                
                myBold = False
                If myTask.OutlineChildren.Count > 0 Then
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
                dateOffset = myD(myTask.Start) - myD(myStartDate)
                myDuration = myD(myTask.Finish) - myD(myTask.Start) + 1
                
                If myTask.Milestone Then
                    
                    mySheet.Cells(currentLine, ganttColumn + dateOffset + 1).Value = ChrW(&H25CA)
                    mySheet.Cells(currentLine, ganttColumn + dateOffset + 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    mySheet.Cells(currentLine, ganttColumn + dateOffset + 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                    
                Else
                    If myTask.OutlineLevel = 1 Then
'  myColor = RGB(48, 84, 150)
' myColor = RGB(66, 96, 162)
                        myColor = RGB(78, 121, 198)
                    ElseIf myTask.OutlineLevel = 2 Then
                        myColor = RGB(142, 169, 219)
                    ElseIf myTask.OutlineLevel = 3 Then
                        myColor = RGB(180, 198, 231)
                    ElseIf myTask.OutlineLevel = 4 Then
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
                   
        End If
                
    Next myTask
                
'Do some grouping =========================================================================================
    currentLine = 4
    For Each myTask In ActiveProject.Tasks
        
        If Not myTask Is Nothing Then
            If myTask.Active Then
                
                currentLine = currentLine + 1
                
                If myTask.OutlineLevel < 8 Then 'levels more than 9 are treated as usual due to MS Excel limitations
                    
                    If myTask.OutlineChildren.Count > 0 Then 'we need to group
                        
                        myStartLine = currentLine
                        myDepth = 0
                        
                        For myIdent = 1 To myTask.OutlineChildren.Count
                            Call myTree(myTask.OutlineChildren(myIdent))
                            Next myIdent
                            
                            myEndLine = myStartLine + myDepth
                            
                            If Not myStartLine = myEndLine Then
                                mySheet.Rows(Trim(Str(myStartLine + 1)) & ":" & Trim(Str(myEndLine))).Group
                                
                            End If
                            
                        End If
                    End If
                    
                    
                End If
            End If
            
            
    Next myTask
'end grouping =============================================================================================
                       
     For i = 1 To ganttColumn
         mySheet.Columns(i).AutoFit
         If mySheet.Columns(i).ColumnWidth > 100 Then
             mySheet.Columns(i).ColumnWidth = 100
         End If
    
     Next i
                            
    For i = ganttColumn + 1 To numberOfDays + ganttColumn + 1
        mySheet.Columns(i).ColumnWidth = 1
    Next i
                                 
   excelapp.ScreenUpdating = True
                               
End Sub
                            
'for gantt chart find minimum and maximum project's start and end dates
Function StartAndEnd()


currentLine = 1

Dim myStartDate As Date
Dim myEndDate As Date

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
            
' some debug info
' Debug.Print "==================="
' Debug.Print "name: """ & myTask.Name & """"
' Debug.Print "unique ID: " & myTask.UniqueID
' Debug.Print "outline number: " & myTask.OutlineNumber
' Debug.Print "outline level: " & myTask.OutlineLevel
' Debug.Print "outline children count: " & myTask.OutlineChildren.Count
            
            
'  some text for the future
'  myTask.ResourceNames
'  myTask.Predecessors
'  myTask.UniqueIDPredecessors
'  myTask.OutlineNumber
'  myTask.Name
'  myTask.Start
'  myTask.StartText
'  myTask.Finish
'  myTask.FinishText
'  myTask.Duration
'  myTask.DurationText
'  myTask.ResourceNames
'  myTask.Predecessors
'  myTask.UniqueIDPredecessors
  
       
        End If

    End If
    
    
Next myTask

    
    Dim res(0 To 1) As Date
    res(0) = myStartDate
    res(1) = myEndDate
    
    StartAndEnd = res
    
End Function

'For those who wants some extention of this module
Function DebugEnum(ByRef myArray)
  '  Dim myArray(1337) As Long
    
    myArray(0) = 205521651    'pjImportResource -- 205521651 -- Represents the Import (Resource) field.
    myArray(1) = 205520915    'pjResourceAccrueAt -- 205520915 -- Represents the Accrue At (Resource) field.
    myArray(2) = 205520907    'pjResourceActualCost -- 205520907 -- Represents the Actual Cost (Resource) field.
    myArray(3) = 205520944    'pjResourceActualOvertimeCost -- 205520944 -- Represents the Actual Overtime Cost (Resource) field.
    myArray(4) = 205520935    'pjResourceActualOvertimeWork -- 205520935 -- Represents the Actual Overtime Work (Resource) field.
    myArray(5) = 205521617    'pjResourceActualOvertimeWorkProtected -- 205521617 -- Represents the Actual Overtime Work Protected (Resource) field.
    myArray(6) = 205520910    'pjResourceActualWork -- 205520910 -- Represents the Actual Work (Resource) field.
    myArray(7) = 205521616    'pjResourceActualWorkProtected -- 205521616 -- Represents the Actual Work Protected (Resource) field.
    myArray(8) = 205520949    'pjResourceACWP -- 205520949 -- Represents the ACWP (Resource) field.
    myArray(9) = 205521153    'pjResourceAssignmentDelay -- 205521153 -- Represents the Assignment Delay (Resource) field.
    myArray(10) = 205521154   'pjResourceAssignmentUnits -- 205521154 -- Represents the Assignment Units (Resource) field.
    myArray(11) = 205520953   'pjResourceAvailableFrom -- 205520953 -- Represents the Available From (Resource) field.
    myArray(12) = 205520954   'pjResourceAvailableTo -- 205520954 -- Represents the Available To (Resource) field.
    myArray(13) = 205520901   'pjResourceBaseCalendar -- 205520901 -- Represents the Base Calendar (Resource) field.
    myArray(14) = 205521693   'pjResourceBaseline10BudgetCost -- 205521693 -- Represents the Baseline10 Budget Cost (Resource) field.
    myArray(15) = 205521692   'pjResourceBaseline10BudgetWork -- 205521692 -- Represents the Baseline10 Budget Work (Resource) field.
    myArray(16) = 205521329   'pjResourceBaseline10Cost -- 205521329 -- Represents the Baseline10 Cost (Resource) field.
    myArray(17) = 205521335   'pjResourceBaseline10Finish -- 205521335 -- Represents the Baseline10 Finish (Resource) field.
    myArray(18) = 205521334   'pjResourceBaseline10Start -- 205521334 -- Represents the Baseline10 Start (Resource) field.
    myArray(19) = 205521328   'pjResourceBaseline10Work -- 205521328 -- Represents the Baseline10 Work (Resource) field.
    myArray(20) = 205521657   'pjResourceBaseline1BudgetCost -- 205521657 -- Represents the Baseline1 Cost Budget (Resource) field.
    myArray(21) = 205521656   'pjResourceBaseline1BudgetWork -- 205521656 -- Represents the Baseline1 Budget Work (Resource) field.
    myArray(22) = 205521239   'pjResourceBaseline1Cost -- 205521239 -- Represents the Baseline1 Cost (Resource) field.
    myArray(23) = 205521245   'pjResourceBaseline1Finish -- 205521245 -- Represents the Baseline1 Finish (Resource) field.
    myArray(24) = 205521244   'pjResourceBaseline1Start -- 205521244 -- Represents the Baseline1 Start (Resource) field.
    myArray(25) = 205521238   'pjResourceBaseline1Work -- 205521238 -- Represents the Baseline1 Work (Resource) field.
    myArray(26) = 205521661   'pjResourceBaseline2BudgetCost -- 205521661 -- Represents the Baseline2 Budget Cost (Resource) field.
    myArray(27) = 205521660   'pjResourceBaseline2BudgetWork -- 205521660 -- Represents the Baseline2 Budget Work (Resource) field.
    myArray(28) = 205521249   'pjResourceBaseline2Cost -- 205521249 -- Represents the Baseline2 Cost (Resource) field.
    myArray(29) = 205521255   'pjResourceBaseline2Finish -- 205521255 -- Represents the Baseline2 Finish (Resource) field.
    myArray(30) = 205521254   'pjResourceBaseline2Start -- 205521254 -- Represents the Baseline2 Start (Resource) field.
    myArray(31) = 205521248   'pjResourceBaseline2Work -- 205521248 -- Represents the Baseline2 Work (Resource) field.
    myArray(32) = 205521665   'pjResourceBaseline3BudgetCost -- 205521665 -- Represents the Baseline3 Cost Budget (Resource) field.
    myArray(33) = 205521664   'pjResourceBaseline3BudgetWork -- 205521664 -- Represents the Baseline3 Budget Work (Resource) field.
    myArray(34) = 205521259   'pjResourceBaseline3Cost -- 205521259 -- Represents the Baseline3 Cost (Resource) field.
    myArray(35) = 205521265   'pjResourceBaseline3Finish -- 205521265 -- Represents the Baseline3 Finish (Resource) field.
    myArray(36) = 205521264   'pjResourceBaseline3Start -- 205521264 -- Represents the Baseline3 Start (Resource) field.
    myArray(37) = 205521258   'pjResourceBaseline3Work -- 205521258 -- Represents the Baseline3 Work (Resource) field.
    myArray(38) = 205521669   'pjResourceBaseline4BudgetCost -- 205521669 -- Represents the Baseline4 Cost Budget (Resource) field.
    myArray(39) = 205521668   'pjResourceBaseline4BudgetWork -- 205521668 -- Represents the Baseline4 Budget Work (Resource) field.
    myArray(40) = 205521269   'pjResourceBaseline4Cost -- 205521269 -- Represents the Baseline4 Cost (Resource) field.
    myArray(41) = 205521275   'pjResourceBaseline4Finish -- 205521275 -- Represents the Baseline4 Finish (Resource) field.
    myArray(42) = 205521274   'pjResourceBaseline4Start -- 205521274 -- Represents the Baseline4 Start (Resource) field.
    myArray(43) = 205521268   'pjResourceBaseline4Work -- 205521268 -- Represents the Baseline5 Work (Resource) field.
    myArray(44) = 205521673   'pjResourceBaseline5BudgetCost -- 205521673 -- Represents the Baseline5 Budget Cost (Resource) field.
    myArray(45) = 205521672   'pjResourceBaseline5BudgetWork -- 205521672 -- Represents the Baseline5 Budget Work (Resource) field.
    myArray(46) = 205521279   'pjResourceBaseline5Cost -- 205521279 -- Represents the Baseline5 Cost (Resource) field.
    myArray(47) = 205521285   'pjResourceBaseline5Finish -- 205521285 -- Represents the Baseline5 Finish (Resource) field.
    myArray(48) = 205521284   'pjResourceBaseline5Start -- 205521284 -- Represents the Baseline5 Start (Resource) field.
    myArray(49) = 205521278   'pjResourceBaseline5Work -- 205521278 -- Represents the Baseline5 Work (Resource) field.
    myArray(50) = 205521677   'pjResourceBaseline6BudgetCost -- 205521677 -- Represents the Baseline6 Budget Cost (Resource) field.
    myArray(51) = 205521676   'pjResourceBaseline6BudgetWork -- 205521676 -- Represents the Baseline6 Budget Work (Resource) field.
    myArray(52) = 205521289   'pjResourceBaseline6Cost -- 205521289 -- Represents the Baseline6 Cost (Resource) field.
    myArray(53) = 205521295   'pjResourceBaseline6Finish -- 205521295 -- Represents the Baseline6 Finish (Resource) field.
    myArray(54) = 205521294   'pjResourceBaseline6Start -- 205521294 -- Represents the Baseline6 Start (Resource) field.
    myArray(55) = 205521288   'pjResourceBaseline6Work -- 205521288 -- Represents the Baseline6 Work (Resource) field.
    myArray(56) = 205521681   'pjResourceBaseline7BudgetCost -- 205521681 -- Represents the Baseline7 Budget Cost (Resource) field.
    myArray(57) = 205521680   'pjResourceBaseline7BudgetWork -- 205521680 -- Represents the Baselin7 Budget Work (Resource) field.
    myArray(58) = 205521299   'pjResourceBaseline7Cost -- 205521299 -- Represents the Baseline7 Cost (Resource) field.
    myArray(59) = 205521305   'pjResourceBaseline7Finish -- 205521305 -- Represents the Baseline7 Finish (Resource) field.
    myArray(60) = 205521304   'pjResourceBaseline7Start -- 205521304 -- Represents the Baseline7 Start (Resource) field.
    myArray(61) = 205521298   'pjResourceBaseline7Work -- 205521298 -- Represents the Baseline7 Work (Resource) field.
    myArray(62) = 205521685   'pjResourceBaseline8BudgetCost -- 205521685 -- Represents the Baseline8 Budget Cost (Resource) field.
    myArray(63) = 205521684   'pjResourceBaseline8BudgetWork -- 205521684 -- Represents the Baseline8 Budget Work (Resource) field.
    myArray(64) = 205521309   'pjResourceBaseline8Cost -- 205521309 -- Represents the Baseline8 Cost (Resource) field.
    myArray(65) = 205521315   'pjResourceBaseline8Finish -- 205521315 -- Represents the Baseline8 Finish (Resource) field.
    myArray(66) = 205521314   'pjResourceBaseline8Start -- 205521314 -- Represents the Baseline8 Start (Resource) field.
    myArray(67) = 205521308   'pjResourceBaseline8Work -- 205521308 -- Represents the Baseline8 Work (Resource) field.
    myArray(68) = 205521689   'pjResourceBaseline9BudgetCost -- 205521689 -- Represents the Baseline9 Budget Cost (Resource) field.
    myArray(69) = 205521688   'pjResourceBaseline9BudgetWork -- 205521688 -- Represents the Baseline9 Budget Work (Resource) field.
    myArray(70) = 205521319   'pjResourceBaseline9Cost -- 205521319 -- Represents the Baseline9 Cost (Resource) field.
    myArray(71) = 205521325   'pjResourceBaseline9Finish -- 205521325 -- Represents the Baseline9 Finish (Resource) field.
    myArray(72) = 205521324   'pjResourceBaseline9Start -- 205521324 -- Represents the Baseline9 Start (Resource) field.
    myArray(73) = 205521318   'pjResourceBaseline9Work -- 205521318 -- Represents the Baseline9 Work (Resource) field.
    myArray(74) = 205521653   'pjResourceBaselineBudgetCost -- 205521653 -- Represents the Baseline Budget Cost (Resource) field.
    myArray(75) = 205521652   'pjResourceBaselineBudgetWork -- 205521652 -- Represents the Baseline Budget Work (Resource) field.
    myArray(76) = 205520913   'pjResourceBaselineCost -- 205520913 -- Represents the Baseline Cost (Resource) field.
    myArray(77) = 205521156   'pjResourceBaselineFinish -- 205521156 -- Represents the Baseline Finish (Resource) field.
    myArray(78) = 205521155   'pjResourceBaselineStart -- 205521155 -- Represents the Baseline Start (Resource) field.
    myArray(79) = 205520911   'pjResourceBaselineWork -- 205520911 -- Represents the Baseline Work (Resource) field.
    myArray(80) = 205520948   'pjResourceBCWP -- 205520948 -- Represents the BCWP (Resource) field.
    myArray(81) = 205520947   'pjResourceBCWS -- 205520947 -- Represents the BCWS (Resource) field.
    myArray(82) = 205521595   'pjResourceBookingType -- 205521595 -- Represents the Booking Type (Resource) field.
    myArray(83) = 205521648   'pjResourceBudget -- 205521648 -- Represents the Budget (Resource) field.
    myArray(84) = 205521650   'pjResourceBudgetCost -- 205521650 -- Represents the Budget Cost (Resource) field.
    myArray(85) = 205521649   'pjResourceBudgetWork -- 205521649 -- Represents the Budget Work (Resource) field.
    myArray(86) = 205521625   'pjResourceCalendarGuid -- 205521625 -- Represents the Calendar GUID (Resource) field.
    myArray(87) = 205521059   'pjResourceCanLevel -- 205521059 -- Represents the Can Level (Resource) field.
    myArray(88) = 205520906   'pjResourceCode -- 205520906 -- Represents the Code (Resource) field.
    myArray(89) = 205521157   'pjResourceConfirmed -- 205521157 -- Represents the Confirmed (Resource) field.
    myArray(90) = 205520908   'pjResourceCost -- 205520908 -- Represents the Cost (Resource) field.
    myArray(91) = 205521019   'pjResourceCost1 -- 205521019 -- Represents the Cost1 (Resource) field.
    myArray(92) = 205521068   'pjResourceCost10 -- 205521068 -- Represents the Cost10 (Resource) field.
    myArray(93) = 205521020   'pjResourceCost2 -- 205521020 -- Represents the Cost2 (Resource) field.
    myArray(94) = 205521021   'pjResourceCost3 -- 205521021 -- Represents the Cost3 (Resource) field.
    myArray(95) = 205521062   'pjResourceCost4 -- 205521062 -- Represents the Cost4 (Resource) field.
    myArray(96) = 205521063   'pjResourceCost5 -- 205521063 -- Represents the Cost5 (Resource) field.
    myArray(97) = 205521064   'pjResourceCost6 -- 205521064 -- Represents the Cost6 (Resource) field.
    myArray(98) = 205521065   'pjResourceCost7 -- 205521065 -- Represents the Cost7 (Resource) field.
    myArray(99) = 205521066   'pjResourceCost8 -- 205521066 -- Represents the Cost8 (Resource) field.
    myArray(100) = 205521067  'pjResourceCost9 -- 205521067 -- Represents the Cost9 (Resource) field.
    myArray(101) = 205521697  'pjResourceCostCenter -- 205521697 -- Represents the Cost Center (Resource) field.
    myArray(102) = 205520914  'pjResourceCostPerUse -- 205520914 -- Represents the Cost Per Use (Resource) field.
    myArray(103) = 205521165  'pjResourceCostRateTable -- 205521165 -- Represents the Cost Rate Table (Resource) field.
    myArray(104) = 205520920  'pjResourceCostVariance -- 205520920 -- Represents the Cost Variance (Resource) field.
    myArray(105) = 205521622  'pjResourceCreated -- 205521622 -- Represents the Created (Resource) field.
    myArray(106) = 205521164  'pjResourceCV -- 205521164 -- Represents the CV (Resource) field.
    myArray(107) = 205521069  'pjResourceDate1 -- 205521069 -- Represents the Date1 (Resource) field.
    myArray(108) = 205521078  'pjResourceDate10 -- 205521078 -- Represents the Date10 (Resource) field.
    myArray(109) = 205521070  'pjResourceDate2 -- 205521070 -- Represents the Date2 (Resource) field.
    myArray(110) = 205521071  'pjResourceDate3 -- 205521071 -- Represents the Date3 (Resource) field.
    myArray(111) = 205521072  'pjResourceDate4 -- 205521072 -- Represents the Date4 (Resource) field.
    myArray(112) = 205521073  'pjResourceDate5 -- 205521073 -- Represents the Date5 (Resource) field.
    myArray(113) = 205521074  'pjResourceDate6 -- 205521074 -- Represents the Date6 (Resource) field.
    myArray(114) = 205521075  'pjResourceDate7 -- 205521075 -- Represents the Date7 (Resource) field.
    myArray(115) = 205521076  'pjResourceDate8 -- 205521076 -- Represents the Date8 (Resource) field.
    myArray(116) = 205521077  'pjResourceDate9 -- 205521077 -- Represents the Date9 (Resource) field.
    myArray(117) = 205521636  'pjResourceDefaultAssignmentOwner -- 205521636 -- Represents the Default Assignment Owner (Resource) field.
    myArray(118) = 205521586  'pjResourceDemandedRequested -- 205521586 -- Represents the Demanded Requested (Resource) field.
    myArray(119) = 205521013  'pjResourceDuration1 -- 205521013 -- Represents the Duration1 (Resource) field.
    myArray(120) = 205521085  'pjResourceDuration10 -- 205521085 -- Represents the Duration10 (Resource) field.
    myArray(121) = 205521014  'pjResourceDuration2 -- 205521014 -- Represents the Duration2 (Resource) field.
    myArray(122) = 205521015  'pjResourceDuration3 -- 205521015 -- Represents the Duration3 (Resource) field.
    myArray(123) = 205521079  'pjResourceDuration4 -- 205521079 -- Represents the Duration4 (Resource) field.
    myArray(124) = 205521080  'pjResourceDuration5 -- 205521080 -- Represents the Duration5 (Resource) field.
    myArray(125) = 205521081  'pjResourceDuration6 -- 205521081 -- Represents the Duration6 (Resource) field.
    myArray(126) = 205521082  'pjResourceDuration7 -- 205521082 -- Represents the Duration7 (Resource) field.
    myArray(127) = 205521083  'pjResourceDuration8 -- 205521083 -- Represents the Duration8 (Resource) field.
    myArray(128) = 205521084  'pjResourceDuration9 -- 205521084 -- Represents the Duration9 (Resource) field.
    myArray(129) = 205520931  'pjResourceEMailAddress -- 205520931 -- Represents the EMail Address (Resource) field.
    myArray(130) = 205521587  'pjResourceEnterprise -- 205521587 -- Represents the Enterprise (Resource) field.
    myArray(131) = 205521583  'pjResourceEnterpriseBaseCalendar -- 205521583 -- Represents the Enterprise Base Calendar (Resource) field.
    myArray(132) = 205521589  'pjResourceEnterpriseCheckedOutBy -- 205521589 -- Represents the Enterprise Checked Out By (Resource) field.
    myArray(133) = 205521342  'pjResourceEnterpriseCost1 -- 205521342 -- Represents the Enterprise Cost1 (Resource) field.
    myArray(134) = 205521351  'pjResourceEnterpriseCost10 -- 205521351 -- Represents the Enterprise Cost10 (Resource) field.
    myArray(135) = 205521343  'pjResourceEnterpriseCost2 -- 205521343 -- Represents the Enterprise Cost2 (Resource) field.
    myArray(136) = 205521344  'pjResourceEnterpriseCost3 -- 205521344 -- Represents the Enterprise Cost3 (Resource) field.
    myArray(137) = 205521345  'pjResourceEnterpriseCost4 -- 205521345 -- Represents the Enterprise Cost4 (Resource) field.
    myArray(138) = 205521346  'pjResourceEnterpriseCost5 -- 205521346 -- Represents the Enterprise Cost5 (Resource) field.
    myArray(139) = 205521347  'pjResourceEnterpriseCost6 -- 205521347 -- Represents the Enterprise Cost6 (Resource) field.
    myArray(140) = 205521348  'pjResourceEnterpriseCost7 -- 205521348 -- Represents the Enterprise Cost7 (Resource) field.
    myArray(141) = 205521349  'pjResourceEnterpriseCost8 -- 205521349 -- Represents the Enterprise Cost8 (Resource) field.
    myArray(142) = 205521350  'pjResourceEnterpriseCost9 -- 205521350 -- Represents the Enterprise Cost9 (Resource) field.
    myArray(143) = 205521352  'pjResourceEnterpriseDate1 -- 205521352 -- Represents the Enterprise Date1 (Resource) field.
    myArray(144) = 205521361  'pjResourceEnterpriseDate10 -- 205521361 -- Represents the Enterprise Date10 (Resource) field.
    myArray(145) = 205521362  'pjResourceEnterpriseDate11 -- 205521362 -- Represents the Enterprise Date11 (Resource) field.
    myArray(146) = 205521363  'pjResourceEnterpriseDate12 -- 205521363 -- Represents the Enterprise Date12 (Resource) field.
    myArray(147) = 205521364  'pjResourceEnterpriseDate13 -- 205521364 -- Represents the Enterprise Date13 (Resource) field.
    myArray(148) = 205521365  'pjResourceEnterpriseDate14 -- 205521365 -- Represents the Enterprise Date14 (Resource) field.
    myArray(149) = 205521366  'pjResourceEnterpriseDate15 -- 205521366 -- Represents the Enterprise Date15 (Resource) field.
    myArray(150) = 205521367  'pjResourceEnterpriseDate16 -- 205521367 -- Represents the Enterprise Date16 (Resource) field.
    myArray(151) = 205521368  'pjResourceEnterpriseDate17 -- 205521368 -- Represents the Enterprise Date17 (Resource) field.
    myArray(152) = 205521369  'pjResourceEnterpriseDate18 -- 205521369 -- Represents the Enterprise Date18 (Resource) field.
    myArray(153) = 205521370  'pjResourceEnterpriseDate19 -- 205521370 -- Represents the Enterprise Date19 (Resource) field.
    myArray(154) = 205521353  'pjResourceEnterpriseDate2 -- 205521353 -- Represents the Enterprise Date2 (Resource) field.
    myArray(155) = 205521371  'pjResourceEnterpriseDate20 -- 205521371 -- Represents the Enterprise Date20 (Resource) field.
    myArray(156) = 205521372  'pjResourceEnterpriseDate21 -- 205521372 -- Represents the Enterprise Date21 (Resource) field.
    myArray(157) = 205521373  'pjResourceEnterpriseDate22 -- 205521373 -- Represents the Enterprise Date22 (Resource) field.
    myArray(158) = 205521374  'pjResourceEnterpriseDate23 -- 205521374 -- Represents the Enterprise Date23 (Resource) field.
    myArray(159) = 205521375  'pjResourceEnterpriseDate24 -- 205521375 -- Represents the Enterprise Date24 (Resource) field.
    myArray(160) = 205521376  'pjResourceEnterpriseDate25 -- 205521376 -- Represents the Enterprise Date25 (Resource) field.
    myArray(161) = 205521377  'pjResourceEnterpriseDate26 -- 205521377 -- Represents the Enterprise Date26 (Resource) field.
    myArray(162) = 205521378  'pjResourceEnterpriseDate27 -- 205521378 -- Represents the Enterprise Date27 (Resource) field.
    myArray(163) = 205521379  'pjResourceEnterpriseDate28 -- 205521379 -- Represents the Enterprise Date28 (Resource) field.
    myArray(164) = 205521380  'pjResourceEnterpriseDate29 -- 205521380 -- Represents the Enterprise Date29 (Resource) field.
    myArray(165) = 205521354  'pjResourceEnterpriseDate3 -- 205521354 -- Represents the Enterprise Date3 (Resource) field.
    myArray(166) = 205521381  'pjResourceEnterpriseDate30 -- 205521381 -- Represents the Enterprise Date30 (Resource) field.
    myArray(167) = 205521355  'pjResourceEnterpriseDate4 -- 205521355 -- Represents the Enterprise Date4 (Resource) field.
    myArray(168) = 205521356  'pjResourceEnterpriseDate5 -- 205521356 -- Represents the Enterprise Date5 (Resource) field.
    myArray(169) = 205521357  'pjResourceEnterpriseDate6 -- 205521357 -- Represents the Enterprise Date6 (Resource) field.
    myArray(170) = 205521358  'pjResourceEnterpriseDate7 -- 205521358 -- Represents the Enterprise Date7 (Resource) field.
    myArray(171) = 205521359  'pjResourceEnterpriseDate8 -- 205521359 -- Represents the Enterprise Date8 (Resource) field.
    myArray(172) = 205521360  'pjResourceEnterpriseDate9 -- 205521360 -- Represents the Enterprise Date9 (Resource) field.
    myArray(173) = 205521382  'pjResourceEnterpriseDuration1 -- 205521382 -- Represents the Enterprise Duration1 (Resource) field.
    myArray(174) = 205521391  'pjResourceEnterpriseDuration10 -- 205521391 -- Represents the Enterprise Duration10 (Resource) field.
    myArray(175) = 205521383  'pjResourceEnterpriseDuration2 -- 205521383 -- Represents the Enterprise Duration2 (Resource) field.
    myArray(176) = 205521384  'pjResourceEnterpriseDuration3 -- 205521384 -- Represents the Enterprise Duration3 (Resource) field.
    myArray(177) = 205521385  'pjResourceEnterpriseDuration4 -- 205521385 -- Represents the Enterprise Duration4 (Resource) field.
    myArray(178) = 205521386  'pjResourceEnterpriseDuration5 -- 205521386 -- Represents the Enterprise Duration5 (Resource) field.
    myArray(179) = 205521387  'pjResourceEnterpriseDuration6 -- 205521387 -- Represents the Enterprise Duration6 (Resource) field.
    myArray(180) = 205521388  'pjResourceEnterpriseDuration7 -- 205521388 -- Represents the Enterprise Duration7 (Resource) field.
    myArray(181) = 205521389  'pjResourceEnterpriseDuration8 -- 205521389 -- Represents the Enterprise Duration8 (Resource) field.
    myArray(182) = 205521390  'pjResourceEnterpriseDuration9 -- 205521390 -- Represents the Enterprise Duration9 (Resource) field.
    myArray(183) = 205521402  'pjResourceEnterpriseFlag1 -- 205521402 -- Represents the Enterprise Flag1 (Resource) field.
    myArray(184) = 205521411  'pjResourceEnterpriseFlag10 -- 205521411 -- Represents the Enterprise Flag10 (Resource) field.
    myArray(185) = 205521412  'pjResourceEnterpriseFlag11 -- 205521412 -- Represents the Enterprise Flag11 (Resource) field.
    myArray(186) = 205521413  'pjResourceEnterpriseFlag12 -- 205521413 -- Represents the Enterprise Flag12 (Resource) field.
    myArray(187) = 205521414  'pjResourceEnterpriseFlag13 -- 205521414 -- Represents the Enterprise Flag13 (Resource) field.
    myArray(188) = 205521415  'pjResourceEnterpriseFlag14 -- 205521415 -- Represents the Enterprise Flag14 (Resource) field.
    myArray(189) = 205521416  'pjResourceEnterpriseFlag15 -- 205521416 -- Represents the Enterprise Flag15 (Resource) field.
    myArray(190) = 205521417  'pjResourceEnterpriseFlag16 -- 205521417 -- Represents the Enterprise Flag16 (Resource) field.
    myArray(191) = 205521418  'pjResourceEnterpriseFlag17 -- 205521418 -- Represents the Enterprise Flag17 (Resource) field.
    myArray(192) = 205521419  'pjResourceEnterpriseFlag18 -- 205521419 -- Represents the Enterprise Flag18 (Resource) field.
    myArray(193) = 205521420  'pjResourceEnterpriseFlag19 -- 205521420 -- Represents the Enterprise Flag19 (Resource) field.
    myArray(194) = 205521403  'pjResourceEnterpriseFlag2 -- 205521403 -- Represents the Enterprise Flag2 (Resource) field.
    myArray(195) = 205521421  'pjResourceEnterpriseFlag20 -- 205521421 -- Represents the Enterprise Flag20 (Resource) field.
    myArray(196) = 205521404  'pjResourceEnterpriseFlag3 -- 205521404 -- Represents the Enterprise Flag3 (Resource) field.
    myArray(197) = 205521405  'pjResourceEnterpriseFlag4 -- 205521405 -- Represents the Enterprise Flag4 (Resource) field.
    myArray(198) = 205521406  'pjResourceEnterpriseFlag5 -- 205521406 -- Represents the Enterprise Flag5 (Resource) field.
    myArray(199) = 205521407  'pjResourceEnterpriseFlag6 -- 205521407 -- Represents the Enterprise Flag6 (Resource) field.
    myArray(200) = 205521408  'pjResourceEnterpriseFlag7 -- 205521408 -- Represents the Enterprise Flag7 (Resource) field.
    myArray(201) = 205521409  'pjResourceEnterpriseFlag8 -- 205521409 -- Represents the Enterprise Flag8 (Resource) field.
    myArray(202) = 205521410  'pjResourceEnterpriseFlag9 -- 205521410 -- Represents the Enterprise Flag9 (Resource) field.
    myArray(203) = 205521582  'pjResourceEnterpriseGeneric -- 205521582 -- Represents the Enterprise Generic (Resource) field.
    myArray(204) = 205521592  'pjResourceEnterpriseInactive -- 205521592 -- Represents the Enterprise Inactive (Resource) field.
    myArray(205) = 205521588  'pjResourceEnterpriseIsCheckedOut -- 205521588 -- Represents the Enterprise Is Checked Out (Resource) field.
    myArray(206) = 205521590  'pjResourceEnterpriseLastModifiedDate -- 205521590 -- Represents the Enterprise Date1 (Resource) field.
    myArray(207) = 205521596  'pjResourceEnterpriseMultiValue20 -- 205521596 -- Represents the Enterprise MultiValue20 (Resource) field.
    myArray(208) = 205521598  'pjResourceEnterpriseMultiValue21 -- 205521598 -- Represents the Enterprise MultiValue21 (Resource) field.
    myArray(209) = 205521600  'pjResourceEnterpriseMultiValue22 -- 205521600 -- Represents the Enterprise MultiValue22 (Resource) field.
    myArray(210) = 205521602  'pjResourceEnterpriseMultiValue23 -- 205521602 -- Represents the Enterprise MultiValue23 (Resource) field.
    myArray(211) = 205521604  'pjResourceEnterpriseMultiValue24 -- 205521604 -- Represents the Enterprise MultiValue24 (Resource) field.
    myArray(212) = 205521606  'pjResourceEnterpriseMultiValue25 -- 205521606 -- Represents the Enterprise MultiValue25 (Resource) field.
    myArray(213) = 205521608  'pjResourceEnterpriseMultiValue26 -- 205521608 -- Represents the Enterprise MultiValue26 (Resource) field.
    myArray(214) = 205521610  'pjResourceEnterpriseMultiValue27 -- 205521610 -- Represents the Enterprise MultiValue27 (Resource) field.
    myArray(215) = 205521612  'pjResourceEnterpriseMultiValue28 -- 205521612 -- Represents the Enterprise MultiValue28 (Resource) field.
    myArray(216) = 205521614  'pjResourceEnterpriseMultiValue29 -- 205521614 -- Represents the Enterprise MultiValue29 (Resource) field.
    myArray(217) = 205521585  'pjResourceEnterpriseNameUsed -- 205521585 -- Represents the Enterprise Name Used (Resource) field.
    myArray(218) = 205521442  'pjResourceEnterpriseNumber1 -- 205521442 -- Represents the Enterprise Number1 (Resource) field.
    myArray(219) = 205521451  'pjResourceEnterpriseNumber10 -- 205521451 -- Represents the Enterprise Number10 (Resource) field.
    myArray(220) = 205521452  'pjResourceEnterpriseNumber11 -- 205521452 -- Represents the Enterprise Number11 (Resource) field.
    myArray(221) = 205521453  'pjResourceEnterpriseNumber12 -- 205521453 -- Represents the Enterprise Number12 (Resource) field.
    myArray(222) = 205521454  'pjResourceEnterpriseNumber13 -- 205521454 -- Represents the Enterprise Number13 (Resource) field.
    myArray(223) = 205521455  'pjResourceEnterpriseNumber14 -- 205521455 -- Represents the Enterprise Number14 (Resource) field.
    myArray(224) = 205521456  'pjResourceEnterpriseNumber15 -- 205521456 -- Represents the Enterprise Number15 (Resource) field.
    myArray(225) = 205521457  'pjResourceEnterpriseNumber16 -- 205521457 -- Represents the Enterprise Number16 (Resource) field.
    myArray(226) = 205521458  'pjResourceEnterpriseNumber17 -- 205521458 -- Represents the Enterprise Number17 (Resource) field.
    myArray(227) = 205521459  'pjResourceEnterpriseNumber18 -- 205521459 -- Represents the Enterprise Number18 (Resource) field.
    myArray(228) = 205521460  'pjResourceEnterpriseNumber19 -- 205521460 -- Represents the Enterprise Number19 (Resource) field.
    myArray(229) = 205521443  'pjResourceEnterpriseNumber2 -- 205521443 -- Represents the Enterprise Number2 (Resource) field.
    myArray(230) = 205521461  'pjResourceEnterpriseNumber20 -- 205521461 -- Represents the Enterprise Number20 (Resource) field.
    myArray(231) = 205521462  'pjResourceEnterpriseNumber21 -- 205521462 -- Represents the Enterprise Number21 (Resource) field.
    myArray(232) = 205521463  'pjResourceEnterpriseNumber22 -- 205521463 -- Represents the Enterprise Number22 (Resource) field.
    myArray(233) = 205521464  'pjResourceEnterpriseNumber23 -- 205521464 -- Represents the Enterprise Number23 (Resource) field.
    myArray(234) = 205521465  'pjResourceEnterpriseNumber24 -- 205521465 -- Represents the Enterprise Number24 (Resource) field.
    myArray(235) = 205521466  'pjResourceEnterpriseNumber25 -- 205521466 -- Represents the Enterprise Number25 (Resource) field.
    myArray(236) = 205521467  'pjResourceEnterpriseNumber26 -- 205521467 -- Represents the Enterprise Number26 (Resource) field.
    myArray(237) = 205521468  'pjResourceEnterpriseNumber27 -- 205521468 -- Represents the Enterprise Number27 (Resource) field.
    myArray(238) = 205521469  'pjResourceEnterpriseNumber28 -- 205521469 -- Represents the Enterprise Number28 (Resource) field.
    myArray(239) = 205521470  'pjResourceEnterpriseNumber29 -- 205521470 -- Represents the Enterprise Number29 (Resource) field.
    myArray(240) = 205521444  'pjResourceEnterpriseNumber3 -- 205521444 -- Represents the Enterprise Number3 (Resource) field.
    myArray(241) = 205521471  'pjResourceEnterpriseNumber30 -- 205521471 -- Represents the Enterprise Number30 (Resource) field.
    myArray(242) = 205521472  'pjResourceEnterpriseNumber31 -- 205521472 -- Represents the Enterprise Number31 (Resource) field.
    myArray(243) = 205521473  'pjResourceEnterpriseNumber32 -- 205521473 -- Represents the Enterprise Number32 (Resource) field.
    myArray(244) = 205521474  'pjResourceEnterpriseNumber33 -- 205521474 -- Represents the Enterprise Number33 (Resource) field.
    myArray(245) = 205521475  'pjResourceEnterpriseNumber34 -- 205521475 -- Represents the Enterprise Number34 (Resource) field.
    myArray(246) = 205521476  'pjResourceEnterpriseNumber35 -- 205521476 -- Represents the Enterprise Number35 (Resource) field.
    myArray(247) = 205521477  'pjResourceEnterpriseNumber36 -- 205521477 -- Represents the Enterprise Number36 (Resource) field.
    myArray(248) = 205521478  'pjResourceEnterpriseNumber37 -- 205521478 -- Represents the Enterprise Number37 (Resource) field.
    myArray(249) = 205521479  'pjResourceEnterpriseNumber38 -- 205521479 -- Represents the Enterprise Number38 (Resource) field.
    myArray(250) = 205521480  'pjResourceEnterpriseNumber39 -- 205521480 -- Represents the Enterprise Number39 (Resource) field.
    myArray(251) = 205521445  'pjResourceEnterpriseNumber4 -- 205521445 -- Represents the Enterprise Number4 (Resource) field.
    myArray(252) = 205521481  'pjResourceEnterpriseNumber40 -- 205521481 -- Represents the Enterprise Number40 (Resource) field.
    myArray(253) = 205521446  'pjResourceEnterpriseNumber5 -- 205521446 -- Represents the Enterprise Number5 (Resource) field.
    myArray(254) = 205521447  'pjResourceEnterpriseNumber6 -- 205521447 -- Represents the Enterprise Number6 (Resource) field.
    myArray(255) = 205521448  'pjResourceEnterpriseNumber7 -- 205521448 -- Represents the Enterprise Number7 (Resource) field.
    myArray(256) = 205521449  'pjResourceEnterpriseNumber8 -- 205521449 -- Represents the Enterprise Number8 (Resource) field.
    myArray(257) = 205521450  'pjResourceEnterpriseNumber9 -- 205521450 -- Represents the Enterprise Number9 (Resource) field.
    myArray(258) = 205521482  'pjResourceEnterpriseOutlineCode1 -- 205521482 -- Represents the Enterprise Outline Code1 (Resource) field.
    myArray(259) = 205521500  'pjResourceEnterpriseOutlineCode10 -- 205521500 -- Represents the Enterprise Outline Code10 (Resource) field.
    myArray(260) = 205521502  'pjResourceEnterpriseOutlineCode11 -- 205521502 -- Represents the Enterprise Outline Code11 (Resource) field.
    myArray(261) = 205521504  'pjResourceEnterpriseOutlineCode12 -- 205521504 -- Represents the Enterprise Outline Code12 (Resource) field.
    myArray(262) = 205521506  'pjResourceEnterpriseOutlineCode13 -- 205521506 -- Represents the Enterprise Outline Code13 (Resource) field.
    myArray(263) = 205521508  'pjResourceEnterpriseOutlineCode14 -- 205521508 -- Represents the Enterprise Outline Code14 (Resource) field.
    myArray(264) = 205521510  'pjResourceEnterpriseOutlineCode15 -- 205521510 -- Represents the Enterprise Outline Code15 (Resource) field.
    myArray(265) = 205521512  'pjResourceEnterpriseOutlineCode16 -- 205521512 -- Represents the Enterprise Outline Code16 (Resource) field.
    myArray(266) = 205521514  'pjResourceEnterpriseOutlineCode17 -- 205521514 -- Represents the Enterprise Outline Code17 (Resource) field.
    myArray(267) = 205521516  'pjResourceEnterpriseOutlineCode18 -- 205521516 -- Represents the Enterprise Outline Code18 (Resource) field.
    myArray(268) = 205521518  'pjResourceEnterpriseOutlineCode19 -- 205521518 -- Represents the Enterprise Outline Code19 (Resource) field.
    myArray(269) = 205521484  'pjResourceEnterpriseOutlineCode2 -- 205521484 -- Represents the Enterprise Outline Code2 (Resource) field.
    myArray(270) = 205521520  'pjResourceEnterpriseOutlineCode20 -- 205521520 -- Represents the Enterprise Outline Code20 (Resource) field.
    myArray(271) = 205521522  'pjResourceEnterpriseOutlineCode21 -- 205521522 -- Represents the Enterprise Outline Code21 (Resource) field.
    myArray(272) = 205521524  'pjResourceEnterpriseOutlineCode22 -- 205521524 -- Represents the Enterprise Outline Code22 (Resource) field.
    myArray(273) = 205521526  'pjResourceEnterpriseOutlineCode23 -- 205521526 -- Represents the Enterprise Outline Code23 (Resource) field.
    myArray(274) = 205521528  'pjResourceEnterpriseOutlineCode24 -- 205521528 -- Represents the Enterprise Outline Code24 (Resource) field.
    myArray(275) = 205521530  'pjResourceEnterpriseOutlineCode25 -- 205521530 -- Represents the Enterprise Outline Code25 (Resource) field.
    myArray(276) = 205521532  'pjResourceEnterpriseOutlineCode26 -- 205521532 -- Represents the Enterprise Outline Code26 (Resource) field.
    myArray(277) = 205521534  'pjResourceEnterpriseOutlineCode27 -- 205521534 -- Represents the Enterprise Outline Code27 (Resource) field.
    myArray(278) = 205521536  'pjResourceEnterpriseOutlineCode28 -- 205521536 -- Represents the Enterprise Outline Code28 (Resource) field.
    myArray(279) = 205521538  'pjResourceEnterpriseOutlineCode29 -- 205521538 -- Represents the Enterprise Outline Code29 (Resource) field.
    myArray(280) = 205521486  'pjResourceEnterpriseOutlineCode3 -- 205521486 -- Represents the Enterprise Outline Code3 (Resource) field.
    myArray(281) = 205521488  'pjResourceEnterpriseOutlineCode4 -- 205521488 -- Represents the Enterprise Outline Code4 (Resource) field.
    myArray(282) = 205521490  'pjResourceEnterpriseOutlineCode5 -- 205521490 -- Represents the Enterprise Outline Code5 (Resource) field.
    myArray(283) = 205521492  'pjResourceEnterpriseOutlineCode6 -- 205521492 -- Represents the Enterprise Outline Code6 (Resource) field.
    myArray(284) = 205521494  'pjResourceEnterpriseOutlineCode7 -- 205521494 -- Represents the Enterprise Outline Code7 (Resource) field.
    myArray(285) = 205521496  'pjResourceEnterpriseOutlineCode8 -- 205521496 -- Represents the Enterprise Outline Code8 (Resource) field.
    myArray(286) = 205521498  'pjResourceEnterpriseOutlineCode9 -- 205521498 -- Represents the Enterprise Outline Code9 (Resource) field.
    myArray(287) = 205521540  'pjResourceEnterpriseRBS -- 205521540 -- Represents the Enterprise RBS (Resource) field.
    myArray(288) = 205521584  'pjResourceEnterpriseRequiredValues -- 205521584 -- Represents the Enterprise Required Values (Resource) field.
    myArray(289) = 205521591  'pjResourceEnterpriseTeamMember -- 205521591 -- Represents the Enterprise Team Member (Resource) field.
    myArray(290) = 205521542  'pjResourceEnterpriseText1 -- 205521542 -- Represents the Enterprise Text1 (Resource) field.
    myArray(291) = 205521551  'pjResourceEnterpriseText10 -- 205521551 -- Represents the Enterprise Text10 (Resource) field.
    myArray(292) = 205521552  'pjResourceEnterpriseText11 -- 205521552 -- Represents the Enterprise Text11 (Resource) field.
    myArray(293) = 205521553  'pjResourceEnterpriseText12 -- 205521553 -- Represents the Enterprise Text12 (Resource) field.
    myArray(294) = 205521554  'pjResourceEnterpriseText13 -- 205521554 -- Represents the Enterprise Text13 (Resource) field.
    myArray(295) = 205521555  'pjResourceEnterpriseText14 -- 205521555 -- Represents the Enterprise Text14 (Resource) field.
    myArray(296) = 205521556  'pjResourceEnterpriseText15 -- 205521556 -- Represents the Enterprise Text15 (Resource) field.
    myArray(297) = 205521557  'pjResourceEnterpriseText16 -- 205521557 -- Represents the Enterprise Text16 (Resource) field.
    myArray(298) = 205521558  'pjResourceEnterpriseText17 -- 205521558 -- Represents the Enterprise Text17 (Resource) field.
    myArray(299) = 205521559  'pjResourceEnterpriseText18 -- 205521559 -- Represents the Enterprise Text18 (Resource) field.
    myArray(300) = 205521560  'pjResourceEnterpriseText19 -- 205521560 -- Represents the Enterprise Text19 (Resource) field.
    myArray(301) = 205521543  'pjResourceEnterpriseText2 -- 205521543 -- Represents the Enterprise Text2 (Resource) field.
    myArray(302) = 205521561  'pjResourceEnterpriseText20 -- 205521561 -- Represents the Enterprise Text20 (Resource) field.
    myArray(303) = 205521562  'pjResourceEnterpriseText21 -- 205521562 -- Represents the Enterprise Text21 (Resource) field.
    myArray(304) = 205521563  'pjResourceEnterpriseText22 -- 205521563 -- Represents the Enterprise Text22 (Resource) field.
    myArray(305) = 205521564  'pjResourceEnterpriseText23 -- 205521564 -- Represents the Enterprise Text23 (Resource) field.
    myArray(306) = 205521565  'pjResourceEnterpriseText24 -- 205521565 -- Represents the Enterprise Text24 (Resource) field.
    myArray(307) = 205521566  'pjResourceEnterpriseText25 -- 205521566 -- Represents the Enterprise Text25 (Resource) field.
    myArray(308) = 205521567  'pjResourceEnterpriseText26 -- 205521567 -- Represents the Enterprise Text26 (Resource) field.
    myArray(309) = 205521568  'pjResourceEnterpriseText27 -- 205521568 -- Represents the Enterprise Text27 (Resource) field.
    myArray(310) = 205521569  'pjResourceEnterpriseText28 -- 205521569 -- Represents the Enterprise Text28 (Resource) field.
    myArray(311) = 205521570  'pjResourceEnterpriseText29 -- 205521570 -- Represents the Enterprise Text29 (Resource) field.
    myArray(312) = 205521544  'pjResourceEnterpriseText3 -- 205521544 -- Represents the Enterprise Text3 (Resource) field.
    myArray(313) = 205521571  'pjResourceEnterpriseText30 -- 205521571 -- Represents the Enterprise Text30 (Resource) field.
    myArray(314) = 205521572  'pjResourceEnterpriseText31 -- 205521572 -- Represents the Enterprise Text31 (Resource) field.
    myArray(315) = 205521573  'pjResourceEnterpriseText32 -- 205521573 -- Represents the Enterprise Text32 (Resource) field.
    myArray(316) = 205521574  'pjResourceEnterpriseText33 -- 205521574 -- Represents the Enterprise Text33 (Resource) field.
    myArray(317) = 205521575  'pjResourceEnterpriseText34 -- 205521575 -- Represents the Enterprise Text34 (Resource) field.
    myArray(318) = 205521576  'pjResourceEnterpriseText35 -- 205521576 -- Represents the Enterprise Text35 (Resource) field.
    myArray(319) = 205521577  'pjResourceEnterpriseText36 -- 205521577 -- Represents the Enterprise Text36 (Resource) field.
    myArray(320) = 205521578  'pjResourceEnterpriseText37 -- 205521578 -- Represents the Enterprise Text37 (Resource) field.
    myArray(321) = 205521579  'pjResourceEnterpriseText38 -- 205521579 -- Represents the Enterprise Text38 (Resource) field.
    myArray(322) = 205521580  'pjResourceEnterpriseText39 -- 205521580 -- Represents the Enterprise Text39 (Resource) field.
    myArray(323) = 205521545  'pjResourceEnterpriseText4 -- 205521545 -- Represents the Enterprise Text4 (Resource) field.
    myArray(324) = 205521581  'pjResourceEnterpriseText40 -- 205521581 -- Represents the Enterprise Text40 (Resource) field.
    myArray(325) = 205521546  'pjResourceEnterpriseText5 -- 205521546 -- Represents the Enterprise Text5 (Resource) field.
    myArray(326) = 205521547  'pjResourceEnterpriseText6 -- 205521547 -- Represents the Enterprise Text6 (Resource) field.
    myArray(327) = 205521548  'pjResourceEnterpriseText7 -- 205521548 -- Represents the Enterprise Text7 (Resource) field.
    myArray(328) = 205521549  'pjResourceEnterpriseText8 -- 205521549 -- Represents the Enterprise Text8 (Resource) field.
    myArray(329) = 205521550  'pjResourceEnterpriseText9 -- 205521550 -- Represents the Enterprise Text9 (Resource) field.
    myArray(330) = 205521339  'pjResourceEnterpriseUniqueID -- 205521339 -- Represents the Enterprise Unique ID (Resource) field.
    myArray(331) = 205521634  'pjResourceErrorMessage -- 205521634 -- Represents the Enterprise Error Message (Resource) field.
    myArray(332) = 205521158  'pjResourceFinish -- 205521158 -- Represents the Finish (Resource) field.
    myArray(333) = 205521003  'pjResourceFinish1 -- 205521003 -- Represents the Finish1 (Resource) field.
    myArray(334) = 205521090  'pjResourceFinish10 -- 205521090 -- Represents the Finish10 (Resource) field.
    myArray(335) = 205521004  'pjResourceFinish2 -- 205521004 -- Represents the Finish2 (Resource) field.
    myArray(336) = 205521005  'pjResourceFinish3 -- 205521005 -- Represents the Finish3 (Resource) field.
    myArray(337) = 205521006  'pjResourceFinish4 -- 205521006 -- Represents the Finish4 (Resource) field.
    myArray(338) = 205521007  'pjResourceFinish5 -- 205521007 -- Represents the Finish5 (Resource) field.
    myArray(339) = 205521086  'pjResourceFinish6 -- 205521086 -- Represents the Finish6 (Resource) field.
    myArray(340) = 205521087  'pjResourceFinish7 -- 205521087 -- Represents the Finish7 (Resource) field.
    myArray(341) = 205521088  'pjResourceFinish8 -- 205521088 -- Represents the Finish8 (Resource) field.
    myArray(342) = 205521089  'pjResourceFinish9 -- 205521089 -- Represents the Finish9 (Resource) field.
    myArray(343) = 205521023  'pjResourceFlag1 -- 205521023 -- Represents the Flag (Resource) field.
    myArray(344) = 205521022  'pjResourceFlag10 -- 205521022 -- Represents the Flag10 (Resource) field.
    myArray(345) = 205521091  'pjResourceFlag11 -- 205521091 -- Represents the Flag11 (Resource) field.
    myArray(346) = 205521092  'pjResourceFlag12 -- 205521092 -- Represents the Flag12 (Resource) field.
    myArray(347) = 205521093  'pjResourceFlag13 -- 205521093 -- Represents the Flag13 (Resource) field.
    myArray(348) = 205521094  'pjResourceFlag14 -- 205521094 -- Represents the Flag14 (Resource) field.
    myArray(349) = 205521095  'pjResourceFlag15 -- 205521095 -- Represents the Flag15 (Resource) field.
    myArray(350) = 205521096  'pjResourceFlag16 -- 205521096 -- Represents the Flag16 (Resource) field.
    myArray(351) = 205521097  'pjResourceFlag17 -- 205521097 -- Represents the Flag17 (Resource) field.
    myArray(352) = 205521098  'pjResourceFlag18 -- 205521098 -- Represents the Flag18 (Resource) field.
    myArray(353) = 205521099  'pjResourceFlag19 -- 205521099 -- Represents the Flag19 (Resource) field.
    myArray(354) = 205521024  'pjResourceFlag2 -- 205521024 -- Represents the Flag2 (Resource) field.
    myArray(355) = 205521100  'pjResourceFlag20 -- 205521100 -- Represents the Flag20 (Resource) field.
    myArray(356) = 205521025  'pjResourceFlag3 -- 205521025 -- Represents the Flag3 (Resource) field.
    myArray(357) = 205521026  'pjResourceFlag4 -- 205521026 -- Represents the Flag4 (Resource) field.
    myArray(358) = 205521027  'pjResourceFlag5 -- 205521027 -- Represents the Flag5 (Resource) field.
    myArray(359) = 205521028  'pjResourceFlag6 -- 205521028 -- Represents the Flag6 (Resource) field.
    myArray(360) = 205521029  'pjResourceFlag7 -- 205521029 -- Represents the Flag7 (Resource) field.
    myArray(361) = 205521030  'pjResourceFlag8 -- 205521030 -- Represents the Flag8 (Resource) field.
    myArray(362) = 205521031  'pjResourceFlag9 -- 205521031 -- Represents the Flag9 (Resource) field.
    myArray(363) = 205520899  'pjResourceGroup -- 205520899 -- Represents the Group (Resource) field.
    myArray(364) = 205521202  'pjResourceGroupbySummary -- 205521202 -- Represents the Group by Summary (Resource) field.
    myArray(365) = 205521624  'pjResourceGuid -- 205521624 -- Represents the Resource GUID (Resource) field.
    myArray(366) = 205521034  'pjResourceHyperlink -- 205521034 -- Represents the Hyperlink (Resource) field.
    myArray(367) = 205521035  'pjResourceHyperlinkAddress -- 205521035 -- Represents the Hyperlink Address (Resource) field.
    myArray(368) = 205521037  'pjResourceHyperlinkHref -- 205521037 -- Represents the Hyperlink HRef (Resource) field.
    myArray(369) = 205521208  'pjResourceHyperlinkScreenTip -- 205521208 -- Represents the Hyperlink ScreenTip (Resource) field.
    myArray(370) = 205521036  'pjResourceHyperlinkSubAddress -- 205521036 -- Represents the Hyperlink Sub Address (Resource) field.
    myArray(371) = 205520896  'pjResourceID -- 205520896 -- Represents the ID (Resource) field.
    myArray(372) = 205521149  'pjResourceIndex -- 205521149 -- Represents the Index (Resource) field.
    myArray(373) = 205520982  'pjResourceIndicators -- 205520982 -- Represents the Indicators (Resource) field.
    myArray(374) = 205520898  'pjResourceInitials -- 205520898 -- Represents the Initials (Resource) field.
    myArray(375) = 205521040  'pjResourceIsAssignment -- 205521040 -- Represents the Is Assignment (Resource) field.
    myArray(376) = 205521696  'pjResourceIsTeam -- 205521696 -- Represents the Is Team (Resource) field.
    myArray(377) = 205521159  'pjResourceLevelingDelay -- 205521159 -- Represents the Leveling Delay (Resource) field.
    myArray(378) = 205520930  'pjResourceLinkedFields -- 205520930 -- Represents the Linked Fields (Resource) field.
    myArray(379) = 205521195  'pjResourceMaterialLabel -- 205521195 -- Represents the Material Label (Resource) field.
    myArray(380) = 205520900  'pjResourceMaxUnits -- 205520900 -- Represents the Max Units (Resource) field.
    myArray(381) = 205520897  'pjResourceName -- 205520897 -- Represents the Name (Resource) field.
    myArray(382) = 205520916  'pjResourceNotes -- 205520916 -- Represents the Notes (Resource) field.
    myArray(383) = 205521008  'pjResourceNumber1 -- 205521008 -- Represents the Number1 (Resource) field.
    myArray(384) = 205521105  'pjResourceNumber10 -- 205521105 -- Represents the Number10 (Resource) field.
    myArray(385) = 205521106  'pjResourceNumber11 -- 205521106 -- Represents the Number11 (Resource) field.
    myArray(386) = 205521107  'pjResourceNumber12 -- 205521107 -- Represents the Number12 (Resource) field.
    myArray(387) = 205521108  'pjResourceNumber13 -- 205521108 -- Represents the Number13 (Resource) field.
    myArray(388) = 205521109  'pjResourceNumber14 -- 205521109 -- Represents the Number14 (Resource) field.
    myArray(389) = 205521110  'pjResourceNumber15 -- 205521110 -- Represents the Number15 (Resource) field.
    myArray(390) = 205521111  'pjResourceNumber16 -- 205521111 -- Represents the Number16 (Resource) field.
    myArray(391) = 205521112  'pjResourceNumber17 -- 205521112 -- Represents the Number17 (Resource) field.
    myArray(392) = 205521113  'pjResourceNumber18 -- 205521113 -- Represents the Number18 (Resource) field.
    myArray(393) = 205521114  'pjResourceNumber19 -- 205521114 -- Represents the Number19 (Resource) field.
    myArray(394) = 205521009  'pjResourceNumber2 -- 205521009 -- Represents the Number2 (Resource) field.
    myArray(395) = 205521115  'pjResourceNumber20 -- 205521115 -- Represents the Number20 (Resource) field.
    myArray(396) = 205521010  'pjResourceNumber3 -- 205521010 -- Represents the Number3 (Resource) field.
    myArray(397) = 205521011  'pjResourceNumber4 -- 205521011 -- Represents the Number4 (Resource) field.
    myArray(398) = 205521012  'pjResourceNumber5 -- 205521012 -- Represents the Number5 (Resource) field.
    myArray(399) = 205521101  'pjResourceNumber6 -- 205521101 -- Represents the Number6 (Resource) field.
    myArray(400) = 205521102  'pjResourceNumber7 -- 205521102 -- Represents the Number7 (Resource) field.
    myArray(401) = 205521103  'pjResourceNumber8 -- 205521103 -- Represents the Number8 (Resource) field.
    myArray(402) = 205521104  'pjResourceNumber9 -- 205521104 -- Represents the Number9 (Resource) field.
    myArray(403) = 205520929  'pjResourceObjects -- 205520929 -- Represents the Objects (Resource) field.
    myArray(404) = 205521174  'pjResourceOutlineCode1 -- 205521174 -- Represents the OutlineCode1 (Resource) field.
    myArray(405) = 205521192  'pjResourceOutlineCode10 -- 205521192 -- Represents the OutlineCode10 (Resource) field.
    myArray(406) = 205521176  'pjResourceOutlineCode2 -- 205521176 -- Represents the OutlineCode2 (Resource) field.
    myArray(407) = 205521178  'pjResourceOutlineCode3 -- 205521178 -- Represents the OutlineCode3(Resource) field.
    myArray(408) = 205521180  'pjResourceOutlineCode4 -- 205521180 -- Represents the OutlineCode4 (Resource) field.
    myArray(409) = 205521182  'pjResourceOutlineCode5 -- 205521182 -- Represents the OutlineCode5 (Resource) field.
    myArray(410) = 205521184  'pjResourceOutlineCode6 -- 205521184 -- Represents the OutlineCode6 (Resource) field.
    myArray(411) = 205521186  'pjResourceOutlineCode7 -- 205521186 -- Represents the OutlineCode7 (Resource) field.
    myArray(412) = 205521188  'pjResourceOutlineCode8 -- 205521188 -- Represents the OutlineCode8 (Resource) field.
    myArray(413) = 205521190  'pjResourceOutlineCode9 -- 205521190 -- Represents the OutlineCode9 (Resource) field.
    myArray(414) = 205520921  'pjResourceOverallocated -- 205520921 -- Represents the Overallocated (Resource) field.
    myArray(415) = 205520943  'pjResourceOvertimeCost -- 205520943 -- Represents the Overtime Cost (Resource) field.
    myArray(416) = 205520903  'pjResourceOvertimeRate -- 205520903 -- Represents the Overtime Rate (Resource) field.
    myArray(417) = 205520912  'pjResourceOvertimeWork -- 205520912 -- Represents the Overtime Work (Resource) field.
    myArray(418) = 205520922  'pjResourcePeakUnits -- 205520922 -- Represents the Peak Units (Resource) field.
    myArray(419) = 205520925  'pjResourcePercentWorkComplete -- 205520925 -- Represents the Percent Work Complete (Resource) field.
    myArray(420) = 205521148  'pjResourcePhonetics -- 205521148 -- Represents the Phonetics (Resource) field.
    myArray(421) = 205521169  'pjResourceProject -- 205521169 -- Represents the Project (Resource) field.
    myArray(422) = 205520934  'pjResourceRegularWork -- 205520934 -- Represents the Regular Work (Resource) field.
    myArray(423) = 205520917  'pjResourceRemainingCost -- 205520917 -- Represents the Remaining Cost (Resource) field.
    myArray(424) = 205520945  'pjResourceRemainingOvertimeCost -- 205520945 -- Represents the Remaining Overtime Cost (Resource) field.
    myArray(425) = 205520936  'pjResourceRemainingOvertimeWork -- 205520936 -- Represents the Remaining Overtime Work (Resource) field.
    myArray(426) = 205520918  'pjResourceRemainingWork -- 205520918 -- Represents the Remaining Work (Resource) field.
    myArray(427) = 205521160  'pjResourceResponsePending -- 205521160 -- Represents the Response Pending (Resource) field.
    myArray(428) = 205520924  'pjResourceSheetNotes -- 205520924 -- Represents the Sheet Notes (Resource) field.
    myArray(429) = 205520902  'pjResourceStandardRate -- 205520902 -- Represents the Standard Rate (Resource) field.
    myArray(430) = 205521161  'pjResourceStart -- 205521161 -- Represents the Start (Resource) field.
    myArray(431) = 205520998  'pjResourceStart1 -- 205520998 -- Represents the Start1 (Resource) field.
    myArray(432) = 205521120  'pjResourceStart10 -- 205521120 -- Represents the Start10 (Resource) field.
    myArray(433) = 205520999  'pjResourceStart2 -- 205520999 -- Represents the Start2 (Resource) field.
    myArray(434) = 205521000  'pjResourceStart3 -- 205521000 -- Represents the Start3 (Resource) field.
    myArray(435) = 205521001  'pjResourceStart4 -- 205521001 -- Represents the Start4 (Resource) field.
    myArray(436) = 205521002  'pjResourceStart5 -- 205521002 -- Represents the Start5 (Resource) field.
    myArray(437) = 205521116  'pjResourceStart6 -- 205521116 -- Represents the Start6 (Resource) field.
    myArray(438) = 205521117  'pjResourceStart7 -- 205521117 -- Represents the Start7 (Resource) field.
    myArray(439) = 205521118  'pjResourceStart8 -- 205521118 -- Represents the Start8 (Resource) field.
    myArray(440) = 205521119  'pjResourceStart9 -- 205521119 -- Represents the Start9 (Resource) field.
    myArray(441) = 205520950  'pjResourceSV -- 205520950 -- Represents the SV (Resource) field.
    myArray(442) = 205521055  'pjResourceTaskSummaryName -- 205521055 -- Represents the Task Summary Name (Resource) field.
    myArray(443) = 205521162  'pjResourceTeamStatusPending -- 205521162 -- Represents the Team Status Pending (Resource) field.
    myArray(444) = 205520904  'pjResourceText1 -- 205520904 -- Represents the Text1 (Resource) field.
    myArray(445) = 205520997  'pjResourceText10 -- 205520997 -- Represents the Text10 (Resource) field.
    myArray(446) = 205521121  'pjResourceText11 -- 205521121 -- Represents the Text11 (Resource) field.
    myArray(447) = 205521122  'pjResourceText12 -- 205521122 -- Represents the Text12 (Resource) field.
    myArray(448) = 205521123  'pjResourceText13 -- 205521123 -- Represents the Text13 (Resource) field.
    myArray(449) = 205521124  'pjResourceText14 -- 205521124 -- Represents the Text14 (Resource) field.
    myArray(450) = 205521125  'pjResourceText15 -- 205521125 -- Represents the Text15 (Resource) field.
    myArray(451) = 205521126  'pjResourceText16 -- 205521126 -- Represents the Text16 (Resource) field.
    myArray(452) = 205521127  'pjResourceText17 -- 205521127 -- Represents the Text17 (Resource) field.
    myArray(453) = 205521128  'pjResourceText18 -- 205521128 -- Represents the Text18 (Resource) field.
    myArray(454) = 205521129  'pjResourceText19 -- 205521129 -- Represents the Text19 (Resource) field.
    myArray(455) = 205520905  'pjResourceText2 -- 205520905 -- Represents the Text2 (Resource) field.
    myArray(456) = 205521130  'pjResourceText20 -- 205521130 -- Represents the Text20 (Resource) field.
    myArray(457) = 205521131  'pjResourceText21 -- 205521131 -- Represents the Text21 (Resource) field.
    myArray(458) = 205521132  'pjResourceText22 -- 205521132 -- Represents the Text22 (Resource) field.
    myArray(459) = 205521133  'pjResourceText23 -- 205521133 -- Represents the Text23 (Resource) field.
    myArray(460) = 205521134  'pjResourceText24 -- 205521134 -- Represents the Text24 (Resource) field.
    myArray(461) = 205521135  'pjResourceText25 -- 205521135 -- Represents the Text25 (Resource) field.
    myArray(462) = 205521136  'pjResourceText26 -- 205521136 -- Represents the Text26 (Resource) field.
    myArray(463) = 205521137  'pjResourceText27 -- 205521137 -- Represents the Text27 (Resource) field.
    myArray(464) = 205521138  'pjResourceText28 -- 205521138 -- Represents the Text28 (Resource) field.
    myArray(465) = 205521139  'pjResourceText29 -- 205521139 -- Represents the Text29 (Resource) field.
    myArray(466) = 205520926  'pjResourceText3 -- 205520926 -- Represents the Text3 (Resource) field.
    myArray(467) = 205521140  'pjResourceText30 -- 205521140 -- Represents the Text30 (Resource) field.
    myArray(468) = 205520927  'pjResourceText4 -- 205520927 -- Represents the Text4 (Resource) field.
    myArray(469) = 205520928  'pjResourceText5 -- 205520928 -- Represents the Text5 (Resource) field.
    myArray(470) = 205520993  'pjResourceText6 -- 205520993 -- Represents the Text6 (Resource) field.
    myArray(471) = 205520994  'pjResourceText7 -- 205520994 -- Represents the Text7 (Resource) field.
    myArray(472) = 205520995  'pjResourceText8 -- 205520995 -- Represents the Text8 (Resource) field.
    myArray(473) = 205520996  'pjResourceText9 -- 205520996 -- Represents the Text9 (Resource) field.
    myArray(474) = 205521196  'pjResourceType -- 205521196 -- Represents the Type (Resource) field.
    myArray(475) = 205520923  'pjResourceUniqueID -- 205520923 -- Represents the UniqueID (Resource) field.
    myArray(476) = 205521163  'pjResourceUpdateNeeded -- 205521163 -- Represents the Update Needed (Resource) field.
    myArray(477) = 205521197  'pjResourceVAC -- 205521197 -- Represents the VAC (Resource) field.
    myArray(478) = 205521236  'pjResourceWBS -- 205521236 -- Represents the WBS (Resource) field.
    myArray(479) = 205521207  'pjResourceWindowsUserAccount -- 205521207 -- Represents the Windows User Account (Resource) field.
    myArray(480) = 205520909  'pjResourceWork -- 205520909 -- Represents the Work (Resource) field.
    myArray(481) = 205521060  'pjResourceWorkContour -- 205521060 -- Represents the Work Contour (Resource) field.
    myArray(482) = 205521168  'pjResourceWorkgroup -- 205521168 -- Represents the Workgroup (Resource) field.
    myArray(483) = 205520919  'pjResourceWorkVariance -- 205520919 -- Represents the Work Variance (Resource) field.
    myArray(484) = 188744959  'pjTaskActive -- 188744959 -- Represents the Active Task (Task) field.
    myArray(485) = 188743687  'pjTaskActualCost -- 188743687 -- Represents the ActualCost (Task) field.
    myArray(486) = 188743708  'pjTaskActualDuration -- 188743708 -- Represents the Actual Duration (Task) field.
    myArray(487) = 188743722  'pjTaskActualFinish -- 188743722 -- Represents the Actual Finish (Task) field.
    myArray(488) = 188743849  'pjTaskActualOvertimeCost -- 188743849 -- Represents the Actual Overtime Cost (Task) field.
    myArray(489) = 188743844  'pjTaskActualOvertimeWork -- 188743844 -- Represents the Actual Overtime Work (Task) field.
    myArray(490) = 188744820  'pjTaskActualOvertimeWorkProtected -- 188744820 -- Represents the Actual Overtime Work Protected (Task) field.
    myArray(491) = 188743721  'pjTaskActualStart -- 188743721 -- Represents the Actual Start (Task) field.
    myArray(492) = 188743682  'pjTaskActualWork -- 188743682 -- Represents the Actual Work (Task) field.
    myArray(493) = 188744819  'pjTaskActualWorkProtected -- 188744819 -- Represents the Actual Work Protected (Task) field.
    myArray(494) = 188743800  'pjTaskACWP -- 188743800 -- Represents the ACWP (Task) field.
    myArray(495) = 188744046  'pjTaskAssignmentDelay -- 188744046 -- Represents the Assignment Delay (Task) field.
    myArray(496) = 188745015  'pjTaskAssignmentPeakUnits -- 188745015 -- Represents the Assignment Peak Units (Task) field.
    myArray(497) = 188744047  'pjTaskAssignmentUnits -- 188744047 -- Represents the Assignment Units (Task) field.
    myArray(498) = 188744927  'pjTaskBaseline10BudgetCost -- 188744927 -- Represents the Baseline10 Budget Cost (Task) field.
    myArray(499) = 188744926  'pjTaskBaseline10BudgetWork -- 188744926 -- Represents the Baseline10 Budget Work (Task) field.
    myArray(500) = 188744270  'pjTaskBaseline10Cost -- 188744270 -- Represents the Baseline10 Cost (Task) field
    myArray(501) = 188744925  'pjTaskBaseline10DeliverableFinish -- 188744925 -- Represents the Baseline10 Deliverable Finish (Task) field.
    myArray(502) = 188744924  'pjTaskBaseline10DeliverableStart -- 188744924 -- Represents the Baseline10 Deliverable Start (Task) field.
    myArray(503) = 188744273  'pjTaskBaseline10Duration -- 188744273 -- Represents the Baseline10 Duration (Task) field.
    myArray(504) = 188744528  'pjTaskBaseline10DurationEstimated -- 188744528 -- Represents the Baseline10 Duration Estimated (Task) field.
    myArray(505) = 188745011  'pjTaskBaseline10DurationText -- 188745011 -- Represents the Baseline10 Duration Text (Task) field.
    myArray(506) = 188744269  'pjTaskBaseline10Finish -- 188744269 -- Represents the Baseline10 Finish (Task) field.
    myArray(507) = 188745010  'pjTaskBaseline10FinishText -- 188745010 -- Represents the Baseline10 Finish Text (Task) field.
    myArray(508) = 188744923  'pjTaskBaseline10FixedCostAccrual -- 188744923 -- Represents the Baseline10 Fixed Cost Accrual (Task) field.
    myArray(509) = 188744268  'pjTaskBaseline10Start -- 188744268 -- Represents the Baseline10 Start (Task) field.
    myArray(510) = 188745009  'pjTaskBaseline10StartText -- 188745009 -- Represents the Baseline10 Start Text (Task) field.
    myArray(511) = 188744271  'pjTaskBaseline10Work -- 188744271 -- Represents the Baseline10 Work (Task) field.
    myArray(512) = 188744864  'pjTaskBaseline1BudgetCost -- 188744864 -- Represents the Baseline1 Budget Cost (Task) field.
    myArray(513) = 188744863  'pjTaskBaseline1BudgetWork -- 188744863 -- Represents the Baseline1 Budget Work (Task) field.
    myArray(514) = 188744164  'pjTaskBaseline1Cost -- 188744164 -- Represents the Baseline1 Cost (Task) field.
    myArray(515) = 188744862  'pjTaskBaseline1DeliverableFinish -- 188744862 -- Represents the Baseline1 Deliverable Finish (Task) field.
    myArray(516) = 188744861  'pjTaskBaseline1DeliverableStart -- 188744861 -- Represents the Baseline1 Deliverable Start (Task) field.
    myArray(517) = 188744167  'pjTaskBaseline1Duration -- 188744167 -- Represents the Baseline1 Duration (Task) field.
    myArray(518) = 188744519  'pjTaskBaseline1DurationEstimated -- 188744519 -- Represents the Baseline1 Duration Estimated (Task) field.
    myArray(519) = 188744984  'pjTaskBaseline1DurationText -- 188744984 -- Represents the Baseline1 Duration Text (Task) field.
    myArray(520) = 188744163  'pjTaskBaseline1Finish -- 188744163 -- Represents the Baseline1 Finish (Task) field.
    myArray(521) = 188744983  'pjTaskBaseline1FinishText -- 188744983 -- Represents the Baseline1 Finish Text (Task) field.
    myArray(522) = 188744860  'pjTaskBaseline1FixedCostAccrual -- 188744860 -- Represents the Baseline1 Fixed Cost Accrual (Task) field.
    myArray(523) = 188744162  'pjTaskBaseline1Start -- 188744162 -- Represents the Baseline1 Start (Task) field.
    myArray(524) = 188744982  'pjTaskBaseline1StartText -- 188744982 -- Represents the Baseline1 Start Text (Task) field.
    myArray(525) = 188744165  'pjTaskBaseline1Work -- 188744165 -- Represents the Baseline1 Budget Work (Task) field.
    myArray(526) = 188744871  'pjTaskBaseline2BudgetCost -- 188744871 -- Represents the Baseline2 Budget Cost (Task) field.
    myArray(527) = 188744870  'pjTaskBaseline2BudgetWork -- 188744870 -- Represents the Baseline2 Budget Work (Task) field.
    myArray(528) = 188744175  'pjTaskBaseline2Cost -- 188744175 -- Represents the Baseline2 Cost (Task) field.
    myArray(529) = 188744869  'pjTaskBaseline2DeliverableFinish -- 188744869 -- Represents the Baseline2 Deliverable Finish (Task) field.
    myArray(530) = 188744868  'pjTaskBaseline2DeliverableStart -- 188744868 -- Represents the Baseline2 Deliverable Start (Task) field.
    myArray(531) = 188744178  'pjTaskBaseline2Duration -- 188744178 -- Represents the Baseline2 Duration (Task) field.
    myArray(532) = 188744520  'pjTaskBaseline2DurationEstimated -- 188744520 -- Represents the Baseline2 Duration Estimated (Task) field.
    myArray(533) = 188744987  'pjTaskBaseline2DurationText -- 188744987 -- Represents the Baseline2 Duration Text (Task) field.
    myArray(534) = 188744174  'pjTaskBaseline2Finish -- 188744174 -- Represents the Baseline2 Finish (Task) field.
    myArray(535) = 188744986  'pjTaskBaseline2FinishText -- 188744986 -- Represents the Baseline2 Finish Text (Task) field.
    myArray(536) = 188744867  'pjTaskBaseline2FixedCostAccrual -- 188744867 -- Represents the Baseline2 Fixed Cost Accrual (Task) field.
    myArray(537) = 188744173  'pjTaskBaseline2Start -- 188744173 -- Represents the Baseline2 Start (Task) field.
    myArray(538) = 188744985  'pjTaskBaseline2StartText -- 188744985 -- Represents the Baseline2 Start Text (Task) field.
    myArray(539) = 188744176  'pjTaskBaseline2Work -- 188744176 -- Represents the Baseline2 Work (Task) field.
    myArray(540) = 188744878  'pjTaskBaseline3BudgetCost -- 188744878 -- Represents the Baseline3 Budget Cost (Task) field.
    myArray(541) = 188744877  'pjTaskBaseline3BudgetWork -- 188744877 -- Represents the Baseline3 Budget Work (Task) field.
    myArray(542) = 188744186  'pjTaskBaseline3Cost -- 188744186 -- Represents the Baseline3 Cost (Task) field.
    myArray(543) = 188744876  'pjTaskBaseline3DeliverableFinish -- 188744876 -- Represents the Baseline3 Deliverable Finish (Task) field.
    myArray(544) = 188744875  'pjTaskBaseline3DeliverableStart -- 188744875 -- Represents the Baseline3 Deliverable Start (Task) field.
    myArray(545) = 188744189  'pjTaskBaseline3Duration -- 188744189 -- Represents the Baseline3 Duration (Task) field.
    myArray(546) = 188744521  'pjTaskBaseline3DurationEstimated -- 188744521 -- Represents the Baseline3 Duration Estimated (Task) field.
    myArray(547) = 188744990  'pjTaskBaseline3DurationText -- 188744990 -- Represents the Baseline3 Duration Text (Task) field.
    myArray(548) = 188744185  'pjTaskBaseline3Finish -- 188744185 -- Represents the Baseline3 Finish (Task) field.
    myArray(549) = 188744989  'pjTaskBaseline3FinishText -- 188744989 -- Represents the Baseline3 Finish Text (Task) field.
    myArray(550) = 188744874  'pjTaskBaseline3FixedCostAccrual -- 188744874 -- Represents the Baseline3 Fixed Cost Accrual (Task) field.
    myArray(551) = 188744184  'pjTaskBaseline3Start -- 188744184 -- Represents the Baseline3 Start (Task) field.
    myArray(552) = 188744988  'pjTaskBaseline3StartText -- 188744988 -- Represents the Baseline3 Start Text (Task) field.
    myArray(553) = 188744187  'pjTaskBaseline3Work -- 188744187 -- Represents the Baseline3 Work (Task) field.
    myArray(554) = 188744885  'pjTaskBaseline4BudgetCost -- 188744885 -- Represents the Baseline4 Budget Cost (Task) field.
    myArray(555) = 188744884  'pjTaskBaseline4BudgetWork -- 188744884 -- Represents the Baseline4 Budget Work (Task) field.
    myArray(556) = 188744197  'pjTaskBaseline4Cost -- 188744197 -- Represents the Baseline4 Cost (Task) field.
    myArray(557) = 188744883  'pjTaskBaseline4DeliverableFinish -- 188744883 -- Represents the Baseline4 Deliverable Finish (Task) field.
    myArray(558) = 188744882  'pjTaskBaseline4DeliverableStart -- 188744882 -- Represents the Baseline4 Deliverable Start (Task) field.
    myArray(559) = 188744200  'pjTaskBaseline4Duration -- 188744200 -- Represents the Baseline4 Duration (Task) field.
    myArray(560) = 188744522  'pjTaskBaseline4DurationEstimated -- 188744522 -- Represents the Baseline4 Duration Estimated (Task) field.
    myArray(561) = 188744993  'pjTaskBaseline4DurationText -- 188744993 -- Represents the Baseline4 Duration Text (Task) field.
    myArray(562) = 188744196  'pjTaskBaseline4Finish -- 188744196 -- Represents the Baseline4 Finish (Task) field.
    myArray(563) = 188744992  'pjTaskBaseline4FinishText -- 188744992 -- Represents the Baseline4 Finish Text (Task) field.
    myArray(564) = 188744881  'pjTaskBaseline4FixedCostAccrual -- 188744881 -- Represents the Baseline4 Fixed Cost Accrual (Task) field
    myArray(565) = 188744195  'pjTaskBaseline4Start -- 188744195 -- Represents the Baseline4 Start (Task) field.
    myArray(566) = 188744991  'pjTaskBaseline4StartText -- 188744991 -- Represents the Baseline4 Start Text (Task) field.
    myArray(567) = 188744198  'pjTaskBaseline4Work -- 188744198 -- Represents the Baseline4 Work (Task) field.
    myArray(568) = 188744892  'pjTaskBaseline5BudgetCost -- 188744892 -- Represents the Baseline5 Budget Cost (Task) field.
    myArray(569) = 188744891  'pjTaskBaseline5BudgetWork -- 188744891 -- Represents the Baseline5 Budget Work (Task) field.
    myArray(570) = 188744208  'pjTaskBaseline5Cost -- 188744208 -- Represents the Baseline5 Cost (Task) field.
    myArray(571) = 188744890  'pjTaskBaseline5DeliverableFinish -- 188744890 -- Represents the Baseline5 Deliverable Finish (Task) field.
    myArray(572) = 188744889  'pjTaskBaseline5DeliverableStart -- 188744889 -- Represents the Baseline5 Deliverable Start (Task) field.
    myArray(573) = 188744211  'pjTaskBaseline5Duration -- 188744211 -- Represents the Baseline5 Duration (Task) field.
    myArray(574) = 188744523  'pjTaskBaseline5DurationEstimated -- 188744523 -- Represents the Baseline5 Duration Estimated (Task) field.
    myArray(575) = 188744996  'pjTaskBaseline5DurationText -- 188744996 -- Represents the Baseline5 Duration Text (Task) field.
    myArray(576) = 188744207  'pjTaskBaseline5Finish -- 188744207 -- Represents the Baseline5 Finish (Task) field.
    myArray(577) = 188744995  'pjTaskBaseline5FinishText -- 188744995 -- Represents the Baseline5 Finish Text (Task) field.
    myArray(578) = 188744888  'pjTaskBaseline5FixedCostAccrual -- 188744888 -- Represents the Baseline5 Fixed Cost Accrual (Task) field.
    myArray(579) = 188744206  'pjTaskBaseline5Start -- 188744206 -- Represents the Baseline5 Start (Task) field.
    myArray(580) = 188744994  'pjTaskBaseline5StartText -- 188744994 -- Represents the Baseline5 Start Text (Task) field.
    myArray(581) = 188744209  'pjTaskBaseline5Work -- 188744209 -- Represents the Baseline5 Work (Task) field.
    myArray(582) = 188744899  'pjTaskBaseline6BudgetCost -- 188744899 -- Represents the Baseline6 Budget Cost (Task) field.
    myArray(583) = 188744898  'pjTaskBaseline6BudgetWork -- 188744898 -- Represents the Baseline6 Budget Work (Task) field.
    myArray(584) = 188744226  'pjTaskBaseline6Cost -- 188744226 -- Represents the Baseline6 Cost (Task) field.
    myArray(585) = 188744897  'pjTaskBaseline6DeliverableFinish -- 188744897 -- Represents the Baseline6 Deliverable Finish (Task) field.
    myArray(586) = 188744896  'pjTaskBaseline6DeliverableStart -- 188744896 -- Represents the Baseline6 Deliverable Start (Task) field.
    myArray(587) = 188744229  'pjTaskBaseline6Duration -- 188744229 -- Represents the Baseline6 Duration (Task) field.
    myArray(588) = 188744524  'pjTaskBaseline6DurationEstimated -- 188744524 -- Represents the Baseline6 Duration Estimated (Task) field.
    myArray(589) = 188744999  'pjTaskBaseline6DurationText -- 188744999 -- Represents the Baseline6 Duration Text (Task) field.
    myArray(590) = 188744225  'pjTaskBaseline6Finish -- 188744225 -- Represents the Baseline6 Finish (Task) field.
    myArray(591) = 188744998  'pjTaskBaseline6FinishText -- 188744998 -- Represents the Baseline6 Finish Text (Task) field.
    myArray(592) = 188744895  'pjTaskBaseline6FixedCostAccrual -- 188744895 -- Represents the Baseline6 Fixed Cost Accrual (Task) field.
    myArray(593) = 188744224  'pjTaskBaseline6Start -- 188744224 -- Represents the Baseline6 Start (Task) field.
    myArray(594) = 188744997  'pjTaskBaseline6StartText -- 188744997 -- Represents the Baseline6 Start Text (Task) field.
    myArray(595) = 188744227  'pjTaskBaseline6Work -- 188744227 -- Represents the Baseline6 Work (Task) field.
    myArray(596) = 188744906  'pjTaskBaseline7BudgetCost -- 188744906 -- Represents the Baseline7 Budget Cost (Task) field.
    myArray(597) = 188744905  'pjTaskBaseline7BudgetWork -- 188744905 -- Represents the Baseline7 Budget Work (Task) field.
    myArray(598) = 188744237  'pjTaskBaseline7Cost -- 188744237 -- Represents the Baseline7 Cost (Task) field.
    myArray(599) = 188744904  'pjTaskBaseline7DeliverableFinish -- 188744904 -- Represents the Baseline7 Deliverable Finish (Task) field.
    myArray(600) = 188744903  'pjTaskBaseline7DeliverableStart -- 188744903 -- Represents the Baseline7 Deliverable Start (Task) field.
    myArray(601) = 188744240  'pjTaskBaseline7Duration -- 188744240 -- Represents the Baseline7 Duration (Task) field.
    myArray(602) = 188744525  'pjTaskBaseline7DurationEstimated -- 188744525 -- Represents the Baseline7 Duration Estimated (Task) field.
    myArray(603) = 188745002  'pjTaskBaseline7DurationText -- 188745002 -- Represents the Baseline7 Duration Text (Task) field.
    myArray(604) = 188744236  'pjTaskBaseline7Finish -- 188744236 -- Represents the Baseline7 Finish (Task) field.
    myArray(605) = 188745001  'pjTaskBaseline7FinishText -- 188745001 -- Represents the Baseline7 Finish Text (Task) field.
    myArray(606) = 188744902  'pjTaskBaseline7FixedCostAccrual -- 188744902 -- Represents the Baseline7 Fixed Cost Accrual (Task) field.
    myArray(607) = 188744235  'pjTaskBaseline7Start -- 188744235 -- Represents the Baseline7 Start (Task) field.
    myArray(608) = 188745000  'pjTaskBaseline7StartText -- 188745000 -- Represents the Baseline7 Start Text (Task) field.
    myArray(609) = 188744238  'pjTaskBaseline7Work -- 188744238 -- Represents the Baseline7 Work (Task) field.
    myArray(610) = 188744913  'pjTaskBaseline8BudgetCost -- 188744913 -- Represents the Baseline8 Budget Cost (Task) field.
    myArray(611) = 188744912  'pjTaskBaseline8BudgetWork -- 188744912 -- Represents the Baseline8 Budget Work (Task) field.
    myArray(612) = 188744248  'pjTaskBaseline8Cost -- 188744248 -- Represents the Baseline8 Cost (Task) field.
    myArray(613) = 188744911  'pjTaskBaseline8DeliverableFinish -- 188744911 -- Represents the Baseline8 Deliverable Finish (Task) field.
    myArray(614) = 188744910  'pjTaskBaseline8DeliverableStart -- 188744910 -- Represents the Baseline8 Deliverable Start (Task) field.
    myArray(615) = 188744251  'pjTaskBaseline8Duration -- 188744251 -- Represents the Baseline8 Duration (Task) field.
    myArray(616) = 188744526  'pjTaskBaseline8DurationEstimated -- 188744526 -- Represents the Baseline8 Duration Estimated (Task) field.
    myArray(617) = 188745005  'pjTaskBaseline8DurationText -- 188745005 -- Represents the Baseline8 Duration Text (Task) field.
    myArray(618) = 188744247  'pjTaskBaseline8Finish -- 188744247 -- Represents the Baseline8 Finish (Task) field.
    myArray(619) = 188745004  'pjTaskBaseline8FinishText -- 188745004 -- Represents the Baseline8 Finish Text (Task) field.
    myArray(620) = 188744909  'pjTaskBaseline8FixedCostAccrual -- 188744909 -- Represents the Baseline8 Fixed Cost Accrual (Task) field.
    myArray(621) = 188744246  'pjTaskBaseline8Start -- 188744246 -- Represents the Baseline8 Start (Task) field.
    myArray(622) = 188745003  'pjTaskBaseline8StartText -- 188745003 -- Represents the Baseline8 Start Text (Task) field.
    myArray(623) = 188744249  'pjTaskBaseline8Work -- 188744249 -- Represents the Baseline8 Work (Task) field.
    myArray(624) = 188744920  'pjTaskBaseline9BudgetCost -- 188744920 -- Represents the Baseline9 Budget Cost (Task) field.
    myArray(625) = 188744919  'pjTaskBaseline9BudgetWork -- 188744919 -- Represents the Baseline9 Budget Work (Task) field.
    myArray(626) = 188744259  'pjTaskBaseline9Cost -- 188744259 -- Represents the Baseline9 Cost (Task) field.
    myArray(627) = 188744918  'pjTaskBaseline9DeliverableFinish -- 188744918 -- Represents the Baseline9 Deliverable Finish (Task) field.
    myArray(628) = 188744917  'pjTaskBaseline9DeliverableStart -- 188744917 -- Represents the Baseline9 Deliverable Start (Task) field.
    myArray(629) = 188744262  'pjTaskBaseline9Duration -- 188744262 -- Represents the Baseline9 Duration (Task) field.
    myArray(630) = 188744527  'pjTaskBaseline9DurationEstimated -- 188744527 -- Represents the Baseline9 Duration Estimated (Task) field.
    myArray(631) = 188745008  'pjTaskBaseline9DurationText -- 188745008 -- Represents the Baseline9 Duration Text (Task) field.
    myArray(632) = 188744258  'pjTaskBaseline9Finish -- 188744258 -- Represents the Baseline9 Finish (Task) field.
    myArray(633) = 188745007  'pjTaskBaseline9FinishText -- 188745007 -- Represents the Baseline9 Finish Text (Task) field.
    myArray(634) = 188744916  'pjTaskBaseline9FixedCostAccrual -- 188744916 -- Represents the Baseline9 Fixed Cost Accrual (Task) field.
    myArray(635) = 188744257  'pjTaskBaseline9Start -- 188744257 -- Represents the Baseline9 Start (Task) field.
    myArray(636) = 188745006  'pjTaskBaseline9StartText -- 188745006 -- Represents the Baseline9 Start Text (Task) field.
    myArray(637) = 188744260  'pjTaskBaseline9Work -- 188744260 -- Represents the Baseline9 Work (Task) field.
    myArray(638) = 188744857  'pjTaskBaselineBudgetCost -- 188744857 -- Represents the Baseline Budget Cost (Task) field.
    myArray(639) = 188744856  'pjTaskBaselineBudgetWork -- 188744856 -- Represents the Baseline Budget Work (Task) field.
    myArray(640) = 188743686  'pjTaskBaselineCost -- 188743686 -- Represents the Baseline Cost (Task) field.
    myArray(641) = 188744855  'pjTaskBaselineDeliverableFinish -- 188744855 -- Represents the Baseline Deliverable Finish (Task) field.
    myArray(642) = 188744854  'pjTaskBaselineDeliverableStart -- 188744854 -- Represents the Baseline Deliverable Start (Task) field.
    myArray(643) = 188743707  'pjTaskBaselineDuration -- 188743707 -- Represents the Baseline Duration (Task) field.
    myArray(644) = 188744093  'pjTaskBaselineDurationEstimated -- 188744093 -- Represents the Baseline Duration Estimated (Task) field.
    myArray(645) = 188744981  'pjTaskBaselineDurationText -- 188744981 -- Represents the Baseline Duration Text (Task) field.
    myArray(646) = 188743724  'pjTaskBaselineFinish -- 188743724 -- Represents the Baseline Finish (Task) field.
    myArray(647) = 188744980  'pjTaskBaselineFinishText -- 188744980 -- Represents the Baseline Finish Text (Task) field.
    myArray(648) = 188744853  'pjTaskBaselineFixedCostAccrual -- 188744853 -- Represents the Baseline Fixed Cost Accrual (Task) field.
    myArray(649) = 188743723  'pjTaskBaselineStart -- 188743723 -- Represents the Baseline Start (Task) field.
    myArray(650) = 188744979  'pjTaskBaselineStartText -- 188744979 -- Represents the Baseline Start Text (Task) field.
    myArray(651) = 188743681  'pjTaskBaselineWork -- 188743681 -- Represents the Baseline Work (Task) field.
    myArray(652) = 188743691  'pjTaskBCWP -- 188743691 -- Represents the BCWP (Task) field.
    myArray(653) = 188743692  'pjTaskBCWS -- 188743692 -- Represents the BCWS (Task) field.
    myArray(654) = 188744852  'pjTaskBudgetCost -- 188744852 -- Represents the Budget Cost (Task) field.
    myArray(655) = 188744851  'pjTaskBudgetWork -- 188744851 -- Represents the Budget Work (Task) field.
    myArray(656) = 188744082  'pjTaskCalendar -- 188744082 -- Represents the Calendar (Task) field.
    myArray(657) = 188744824  'pjTaskCalendarGuid -- 188744824 -- Represents the Calendar GUID (Task) field.
    myArray(658) = 188743790  'pjTaskConfirmed -- 188743790 -- Represents the Confirmed (Task) field.
    myArray(659) = 188743698  'pjTaskConstraintDate -- 188743698 -- Represents the Constraint Date (Task) field.
    myArray(660) = 188743697  'pjTaskConstraintType -- 188743697 -- Represents the Constraint Type (Task) field.
    myArray(661) = 188743792  'pjTaskContact -- 188743792 -- Represents the Contact (Task) field.
    myArray(662) = 188743685  'pjTaskCost -- 188743685 -- Represents the Cost (Task) field.
    myArray(663) = 188743786  'pjTaskCost1 -- 188743786 -- Represents the Cost1 (Task) field.
    myArray(664) = 188743944  'pjTaskCost10 -- 188743944 -- Represents the Cost10 (Task) field.
    myArray(665) = 188743787  'pjTaskCost2 -- 188743787 -- Represents the Cost2 (Task) field.
    myArray(666) = 188743788  'pjTaskCost3 -- 188743788 -- Represents the Cost3 (Task) field.
    myArray(667) = 188743938  'pjTaskCost4 -- 188743938 -- Represents the Cost4 (Task) field.
    myArray(668) = 188743939  'pjTaskCost5 -- 188743939 -- Represents the Cost5 (Task) field.
    myArray(669) = 188743940  'pjTaskCost6 -- 188743940 -- Represents the Cost6 (Task) field.
    myArray(670) = 188743941  'pjTaskCost7 -- 188743941 -- Represents the Cost7 (Task) field.
    myArray(671) = 188743942  'pjTaskCost8 -- 188743942 -- Represents the Cost8 (Task) field.
    myArray(672) = 188743943  'pjTaskCost9 -- 188743943 -- Represents the Cost9 (Task) field.
    myArray(673) = 188744048  'pjTaskCostRateTable -- 188744048 -- Represents the Cost Rate Table (Task) field.
    myArray(674) = 188743689  'pjTaskCostVariance -- 188743689 -- Represents the Cost Variance (Task) field.
    myArray(675) = 188744217  'pjTaskCPI -- 188744217 -- Represents the CPI (Task) field.
    myArray(676) = 188743773  'pjTaskCreated -- 188743773 -- Represents the Created (Task) field.
    myArray(677) = 188743699  'pjTaskCritical -- 188743699 -- Represents the Critical (Task) field.
    myArray(678) = 188743763  'pjTaskCV -- 188743763 -- Represents the CV (Task) field.
    myArray(679) = 188744219  'pjTaskCVPercent -- 188744219 -- Represents the CV Percent (Task) field.
    myArray(680) = 188743945  'pjTaskDate1 -- 188743945 -- Represents the Date1 (Task) field.
    myArray(681) = 188743954  'pjTaskDate10 -- 188743954 -- Represents the Date10 (Task) field.
    myArray(682) = 188743946  'pjTaskDate2 -- 188743946 -- Represents the Date2 (Task) field.
    myArray(683) = 188743947  'pjTaskDate3 -- 188743947 -- Represents the Date3 (Task) field.
    myArray(684) = 188743948  'pjTaskDate4 -- 188743948 -- Represents the Date4 (Task) field.
    myArray(685) = 188743949  'pjTaskDate5 -- 188743949 -- Represents the Date5 (Task) field.
    myArray(686) = 188743950  'pjTaskDate6 -- 188743950 -- Represents the Date6 (Task) field.
    myArray(687) = 188743951  'pjTaskDate7 -- 188743951 -- Represents the Date7 (Task) field.
    myArray(688) = 188743952  'pjTaskDate8 -- 188743952 -- Represents the Date8 (Task) field.
    myArray(689) = 188743953  'pjTaskDate9 -- 188743953 -- Represents the Date9 (Task) field.
    myArray(690) = 188744117  'pjTaskDeadline -- 188744117 -- Represents the Deadline (Task) field.
    myArray(691) = 188743700  'pjTaskDelay -- 188743700 -- Represents the Delay (Task) field.
    myArray(692) = 188744833  'pjTaskDeliverableFinish -- 188744833 -- Represents the Deliverable Finish (Task) field.
    myArray(693) = 188744826  'pjTaskDeliverableGuid -- 188744826 -- Represents the Deliverable GUID (Task) field.
    myArray(694) = 188744956  'pjTaskDeliverableName -- 188744956 -- Represents the Deliverable Start (Task) field.
    myArray(695) = 188744832  'pjTaskDeliverableStart -- 188744832 -- Represents the Deliverable Start (Task) field.
    myArray(696) = 188744827  'pjTaskDeliverableType -- 188744827 -- Represents the Deliverable Type (Task) field.
    myArray(697) = 188744800  'pjTaskDemandedRequested -- 188744800 -- Represents the Demanded Requested (Task) field.
    myArray(698) = 188743709  'pjTaskDuration -- 188743709 -- Represents the Duration (Task) field.
    myArray(699) = 188743783  'pjTaskDuration1 -- 188743783 -- Represents the Duration1 (Task) field.
    myArray(700) = 188743961  'pjTaskDuration10 -- 188743961 -- Represents the Duration10 (Task) field.
    myArray(701) = 188744092  'pjTaskDuration10Estimated -- 188744092 -- Represents the Duration10 Estimated (Task) field.
    myArray(702) = 188744083  'pjTaskDuration1Estimated -- 188744083 -- Represents the Duration11 Estimated (Task) field.
    myArray(703) = 188743784  'pjTaskDuration2 -- 188743784 -- Represents the Duration2 (Task) field.
    myArray(704) = 188744084  'pjTaskDuration2Estimated -- 188744084 -- Represents the Duration2 Estimated (Task) field.
    myArray(705) = 188743785  'pjTaskDuration3 -- 188743785 -- Represents the Duration3 (Task) field.
    myArray(706) = 188744085  'pjTaskDuration3Estimated -- 188744085 -- Represents the Duration3 Estimated (Task) field.
    myArray(707) = 188743955  'pjTaskDuration4 -- 188743955 -- Represents the Duration4 (Task) field.
    myArray(708) = 188744086  'pjTaskDuration4Estimated -- 188744086 -- Represents the Duration4 Estimated (Task) field.
    myArray(709) = 188743956  'pjTaskDuration5 -- 188743956 -- Represents the Duration5 (Task) field.
    myArray(710) = 188744087  'pjTaskDuration5Estimated -- 188744087 -- Represents the Duration5 (Estimated Task) field.
    myArray(711) = 188743957  'pjTaskDuration6 -- 188743957 -- Represents the Duration6 (Task) field.
    myArray(712) = 188744088  'pjTaskDuration6Estimated -- 188744088 -- Represents the Duration6 Estimated (Task) field.
    myArray(713) = 188743958  'pjTaskDuration7 -- 188743958 -- Represents the Duration7 (Task) field.
    myArray(714) = 188744089  'pjTaskDuration7Estimated -- 188744089 -- Represents the Duration7 Estimated (Task) field.
    myArray(715) = 188743959  'pjTaskDuration8 -- 188743959 -- Represents the Duration8 (Task) field.
    myArray(716) = 188744090  'pjTaskDuration8Estimated -- 188744090 -- Represents the Duration8 Estimated (Task) field.
    myArray(717) = 188743960  'pjTaskDuration9 -- 188743960 -- Represents the Duration9 (Task) field.
    myArray(718) = 188744091  'pjTaskDuration9Estimated -- 188744091 -- Represents the Duration9 Estimated (Task) field.
    myArray(719) = 188744967  'pjTaskDurationText -- 188744967 -- Represents the Task Duration Text (Task) field.
    myArray(720) = 188743710  'pjTaskDurationVariance -- 188743710 -- Represents the Duration Variance (Task) field.
    myArray(721) = 188744221  'pjTaskEAC -- 188744221 -- Represents the EAC (Task) field.
    myArray(722) = 188743718  'pjTaskEarlyFinish -- 188743718 -- Represents the Early Finish (Task) field.
    myArray(723) = 188743717  'pjTaskEarlyStart -- 188743717 -- Represents the Early Start (Task) field.
    myArray(724) = 188744802  'pjTaskEarnedValueMethod -- 188744802 -- Represents the Earned Value Method (Task) field.
    myArray(725) = 188743812  'pjTaskEffortDriven -- 188743812 -- Represents the Effort Driven (Task) field.
    myArray(726) = 188744279  'pjTaskEnterpriseCost1 -- 188744279 -- Represents the Enterprise Cost1 (Task) field.
    myArray(727) = 188744288  'pjTaskEnterpriseCost10 -- 188744288 -- Represents the Enterprise Cost10 (Task) field.
    myArray(728) = 188744280  'pjTaskEnterpriseCost2 -- 188744280 -- Represents the Enterprise Cost2 (Task) field.
    myArray(729) = 188744281  'pjTaskEnterpriseCost3 -- 188744281 -- Represents the Enterprise Cost3 (Task) field.
    myArray(730) = 188744282  'pjTaskEnterpriseCost4 -- 188744282 -- Represents the Enterprise Cost4 (Task) field.
    myArray(731) = 188744283  'pjTaskEnterpriseCost5 -- 188744283 -- Represents the Enterprise Cost5 (Task) field.
    myArray(732) = 188744284  'pjTaskEnterpriseCost6 -- 188744284 -- Represents the Enterprise Cost6 (Task) field.
    myArray(733) = 188744285  'pjTaskEnterpriseCost7 -- 188744285 -- Represents the Enterprise Cost7 (Task) field.
    myArray(734) = 188744286  'pjTaskEnterpriseCost8 -- 188744286 -- Represents the Enterprise Cost8 (Task) field.
    myArray(735) = 188744287  'pjTaskEnterpriseCost9 -- 188744287 -- Represents the Enterprise Cost9 (Task) field.
    myArray(736) = 188744289  'pjTaskEnterpriseDate1 -- 188744289 -- Represents the Enterprise Date1 (Task) field.
    myArray(737) = 188744298  'pjTaskEnterpriseDate10 -- 188744298 -- Represents the Enterprise Date10 (Task) field.
    myArray(738) = 188744299  'pjTaskEnterpriseDate11 -- 188744299 -- Represents the Enterprise Date11 (Task) field.
    myArray(739) = 188744300  'pjTaskEnterpriseDate12 -- 188744300 -- Represents the Enterprise Date12 (Task) field.
    myArray(740) = 188744301  'pjTaskEnterpriseDate13 -- 188744301 -- Represents the Enterprise Date13 (Task) field.
    myArray(741) = 188744302  'pjTaskEnterpriseDate14 -- 188744302 -- Represents the Enterprise Date14 (Task) field.
    myArray(742) = 188744303  'pjTaskEnterpriseDate15 -- 188744303 -- Represents the Enterprise Date15 (Task) field.
    myArray(743) = 188744304  'pjTaskEnterpriseDate16 -- 188744304 -- Represents the Enterprise Date16 (Task) field.
    myArray(744) = 188744305  'pjTaskEnterpriseDate17 -- 188744305 -- Represents the Enterprise Date17 (Task) field.
    myArray(745) = 188744306  'pjTaskEnterpriseDate18 -- 188744306 -- Represents the Enterprise Date18 (Task) field.
    myArray(746) = 188744307  'pjTaskEnterpriseDate19 -- 188744307 -- Represents the Enterprise Date19 (Task) field.
    myArray(747) = 188744290  'pjTaskEnterpriseDate2 -- 188744290 -- Represents the Enterprise Date2(Task) field.
    myArray(748) = 188744308  'pjTaskEnterpriseDate20 -- 188744308 -- Represents the Enterprise Date20 (Task) field.
    myArray(749) = 188744309  'pjTaskEnterpriseDate21 -- 188744309 -- Represents the Enterprise Date21 (Task) field.
    myArray(750) = 188744310  'pjTaskEnterpriseDate22 -- 188744310 -- Represents the Enterprise Date22 (Task) field.
    myArray(751) = 188744311  'pjTaskEnterpriseDate23 -- 188744311 -- Represents the Enterprise Date23 (Task) field.
    myArray(752) = 188744312  'pjTaskEnterpriseDate24 -- 188744312 -- Represents the Enterprise Date24 (Task) field.
    myArray(753) = 188744313  'pjTaskEnterpriseDate25 -- 188744313 -- Represents the Enterprise Date25 (Task) field.
    myArray(754) = 188744314  'pjTaskEnterpriseDate26 -- 188744314 -- Represents the Enterprise Date26 (Task) field.
    myArray(755) = 188744315  'pjTaskEnterpriseDate27 -- 188744315 -- Represents the Enterprise Date27 (Task) field.
    myArray(756) = 188744316  'pjTaskEnterpriseDate28 -- 188744316 -- Represents the Enterprise Date28 (Task) field.
    myArray(757) = 188744317  'pjTaskEnterpriseDate29 -- 188744317 -- Represents the Enterprise Date29 (Task) field.
    myArray(758) = 188744291  'pjTaskEnterpriseDate3 -- 188744291 -- Represents the Enterprise Date3 (Task) field.
    myArray(759) = 188744318  'pjTaskEnterpriseDate30 -- 188744318 -- Represents the Enterprise Date30 (Task) field.
    myArray(760) = 188744292  'pjTaskEnterpriseDate4 -- 188744292 -- Represents the Enterprise Date4 (Task) field.
    myArray(761) = 188744293  'pjTaskEnterpriseDate5 -- 188744293 -- Represents the Enterprise Date5 (Task) field.
    myArray(762) = 188744294  'pjTaskEnterpriseDate6 -- 188744294 -- Represents the Enterprise Date6 (Task) field.
    myArray(763) = 188744295  'pjTaskEnterpriseDate7 -- 188744295 -- Represents the Enterprise Date7 (Task) field.
    myArray(764) = 188744296  'pjTaskEnterpriseDate8 -- 188744296 -- Represents the Enterprise Date8 (Task) field.
    myArray(765) = 188744297  'pjTaskEnterpriseDate9 -- 188744297 -- Represents the Enterprise Date9 (Task) field.
    myArray(766) = 188744319  'pjTaskEnterpriseDuration1 -- 188744319 -- Represents the Enterprise Duration1 (Task) field.
    myArray(767) = 188744328  'pjTaskEnterpriseDuration10 -- 188744328 -- Represents the Enterprise Duration10 (Task) field.
    myArray(768) = 188744320  'pjTaskEnterpriseDuration2 -- 188744320 -- Represents the Enterprise Duration2 (Task) field.
    myArray(769) = 188744321  'pjTaskEnterpriseDuration3 -- 188744321 -- Represents the Enterprise Duration3 (Task) field.
    myArray(770) = 188744322  'pjTaskEnterpriseDuration4 -- 188744322 -- Represents the Enterprise Duration4 (Task) field.
    myArray(771) = 188744323  'pjTaskEnterpriseDuration5 -- 188744323 -- Represents the Enterprise Duration5 (Task) field.
    myArray(772) = 188744324  'pjTaskEnterpriseDuration6 -- 188744324 -- Represents the Enterprise Duration6 (Task) field.
    myArray(773) = 188744325  'pjTaskEnterpriseDuration7 -- 188744325 -- Represents the Enterprise Duration7 (Task) field.
    myArray(774) = 188744326  'pjTaskEnterpriseDuration8 -- 188744326 -- Represents the Enterprise Duration8 (Task) field.
    myArray(775) = 188744327  'pjTaskEnterpriseDuration9 -- 188744327 -- Represents the Enterprise Duration9 (Task) field.
    myArray(776) = 188744339  'pjTaskEnterpriseFlag1 -- 188744339 -- Represents the Enterprise Flag1 (Task) field.
    myArray(777) = 188744348  'pjTaskEnterpriseFlag10 -- 188744348 -- Represents the Enterprise Flag10 (Task) field.
    myArray(778) = 188744349  'pjTaskEnterpriseFlag11 -- 188744349 -- Represents the Enterprise Flag11 (Task) field.
    myArray(779) = 188744350  'pjTaskEnterpriseFlag12 -- 188744350 -- Represents the Enterprise Flag12 (Task) field.
    myArray(780) = 188744351  'pjTaskEnterpriseFlag13 -- 188744351 -- Represents the Enterprise Flag13 (Task) field.
    myArray(781) = 188744352  'pjTaskEnterpriseFlag14 -- 188744352 -- Represents the Enterprise Flag14 (Task) field.
    myArray(782) = 188744353  'pjTaskEnterpriseFlag15 -- 188744353 -- Represents the Enterprise Flag15 (Task) field.
    myArray(783) = 188744354  'pjTaskEnterpriseFlag16 -- 188744354 -- Represents the Enterprise Flag16 (Task) field.
    myArray(784) = 188744355  'pjTaskEnterpriseFlag17 -- 188744355 -- Represents the Enterprise Flag17 (Task) field.
    myArray(785) = 188744356  'pjTaskEnterpriseFlag18 -- 188744356 -- Represents the Enterprise Flag18 (Task) field.
    myArray(786) = 188744357  'pjTaskEnterpriseFlag19 -- 188744357 -- Represents the Enterprise Flag19 (Task) field.
    myArray(787) = 188744340  'pjTaskEnterpriseFlag2 -- 188744340 -- Represents the Enterprise Flag2 (Task) field.
    myArray(788) = 188744358  'pjTaskEnterpriseFlag20 -- 188744358 -- Represents the Enterprise Flag20 (Task) field.
    myArray(789) = 188744341  'pjTaskEnterpriseFlag3 -- 188744341 -- Represents the Enterprise Flag3 (Task) field.
    myArray(790) = 188744342  'pjTaskEnterpriseFlag4 -- 188744342 -- Represents the Enterprise Flag4 (Task) field.
    myArray(791) = 188744343  'pjTaskEnterpriseFlag5 -- 188744343 -- Represents the Enterprise Flag5 (Task) field.
    myArray(792) = 188744344  'pjTaskEnterpriseFlag6 -- 188744344 -- Represents the Enterprise Flag6 (Task) field.
    myArray(793) = 188744345  'pjTaskEnterpriseFlag7 -- 188744345 -- Represents the Enterprise Flag7 (Task) field.
    myArray(794) = 188744346  'pjTaskEnterpriseFlag8 -- 188744346 -- Represents the Enterprise Flag8 (Task) field.
    myArray(795) = 188744347  'pjTaskEnterpriseFlag9 -- 188744347 -- Represents the Enterprise Flag9 (Task) field.
    myArray(796) = 188744379  'pjTaskEnterpriseNumber1 -- 188744379 -- Represents the Enterprise Number1 (Task) field.
    myArray(797) = 188744388  'pjTaskEnterpriseNumber10 -- 188744388 -- Represents the Enterprise Number10 (Task) field.
    myArray(798) = 188744389  'pjTaskEnterpriseNumber11 -- 188744389 -- Represents the Enterprise Number11 (Task) field.
    myArray(799) = 188744390  'pjTaskEnterpriseNumber12 -- 188744390 -- Represents the Enterprise Number12 (Task) field.
    myArray(800) = 188744391  'pjTaskEnterpriseNumber13 -- 188744391 -- Represents the Enterprise Number13 (Task) field.
    myArray(801) = 188744392  'pjTaskEnterpriseNumber14 -- 188744392 -- Represents the Enterprise Number14 (Task) field.
    myArray(802) = 188744393  'pjTaskEnterpriseNumber15 -- 188744393 -- Represents the Enterprise Number15 (Task) field.
    myArray(803) = 188744394  'pjTaskEnterpriseNumber16 -- 188744394 -- Represents the Enterprise Number16 (Task) field.
    myArray(804) = 188744395  'pjTaskEnterpriseNumber17 -- 188744395 -- Represents the Enterprise Number17 (Task) field.
    myArray(805) = 188744396  'pjTaskEnterpriseNumber18 -- 188744396 -- Represents the Enterprise Number18 (Task) field.
    myArray(806) = 188744397  'pjTaskEnterpriseNumber19 -- 188744397 -- Represents the Enterprise Number19 (Task) field.
    myArray(807) = 188744380  'pjTaskEnterpriseNumber2 -- 188744380 -- Represents the Enterprise Number2 (Task) field.
    myArray(808) = 188744398  'pjTaskEnterpriseNumber20 -- 188744398 -- Represents the Enterprise Number20 (Task) field.
    myArray(809) = 188744399  'pjTaskEnterpriseNumber21 -- 188744399 -- Represents the Enterprise Number21 (Task) field.
    myArray(810) = 188744400  'pjTaskEnterpriseNumber22 -- 188744400 -- Represents the Enterprise Number22 (Task) field.
    myArray(811) = 188744401  'pjTaskEnterpriseNumber23 -- 188744401 -- Represents the Enterprise Number23 (Task) field.
    myArray(812) = 188744402  'pjTaskEnterpriseNumber24 -- 188744402 -- Represents the Enterprise Number24 (Task) field.
    myArray(813) = 188744403  'pjTaskEnterpriseNumber25 -- 188744403 -- Represents the Enterprise Number25 (Task) field.
    myArray(814) = 188744404  'pjTaskEnterpriseNumber26 -- 188744404 -- Represents the Enterprise Number26 (Task) field.
    myArray(815) = 188744405  'pjTaskEnterpriseNumber27 -- 188744405 -- Represents the Enterprise Number27 (Task) field.
    myArray(816) = 188744406  'pjTaskEnterpriseNumber28 -- 188744406 -- Represents the Enterprise Number28 (Task) field.
    myArray(817) = 188744407  'pjTaskEnterpriseNumber29 -- 188744407 -- Represents the Enterprise Number29 (Task) field.
    myArray(818) = 188744381  'pjTaskEnterpriseNumber3 -- 188744381 -- Represents the Enterprise Number3 (Task) field.
    myArray(819) = 188744408  'pjTaskEnterpriseNumber30 -- 188744408 -- Represents the Enterprise Number30 (Task) field.
    myArray(820) = 188744409  'pjTaskEnterpriseNumber31 -- 188744409 -- Represents the Enterprise Number31 (Task) field.
    myArray(821) = 188744410  'pjTaskEnterpriseNumber32 -- 188744410 -- Represents the Enterprise Number32 (Task) field.
    myArray(822) = 188744411  'pjTaskEnterpriseNumber33 -- 188744411 -- Represents the Enterprise Number33 (Task) field.
    myArray(823) = 188744412  'pjTaskEnterpriseNumber34 -- 188744412 -- Represents the Enterprise Number34 (Task) field.
    myArray(824) = 188744413  'pjTaskEnterpriseNumber35 -- 188744413 -- Represents the Enterprise Number35 (Task) field.
    myArray(825) = 188744414  'pjTaskEnterpriseNumber36 -- 188744414 -- Represents the Enterprise Number36 (Task) field.
    myArray(826) = 188744415  'pjTaskEnterpriseNumber37 -- 188744415 -- Represents the Enterprise Number37 (Task) field.
    myArray(827) = 188744416  'pjTaskEnterpriseNumber38 -- 188744416 -- Represents the Enterprise Number38 (Task) field.
    myArray(828) = 188744417  'pjTaskEnterpriseNumber39 -- 188744417 -- Represents the Enterprise Number39 (Task) field.
    myArray(829) = 188744382  'pjTaskEnterpriseNumber4 -- 188744382 -- Represents the Enterprise Number4 (Task) field.
    myArray(830) = 188744418  'pjTaskEnterpriseNumber40 -- 188744418 -- Represents the Enterprise Number40 (Task) field.
    myArray(831) = 188744383  'pjTaskEnterpriseNumber5 -- 188744383 -- Represents the Enterprise Number5 (Task) field.
    myArray(832) = 188744384  'pjTaskEnterpriseNumber6 -- 188744384 -- Represents the Enterprise Number6 (Task) field.
    myArray(833) = 188744385  'pjTaskEnterpriseNumber7 -- 188744385 -- Represents the Enterprise Number7 (Task) field.
    myArray(834) = 188744386  'pjTaskEnterpriseNumber8 -- 188744386 -- Represents the Enterprise Number8 (Task) field.
    myArray(835) = 188744387  'pjTaskEnterpriseNumber9 -- 188744387 -- Represents the Enterprise Number9 (Task) field.
    myArray(836) = 188744419  'pjTaskEnterpriseOutlineCode1 -- 188744419 -- Represents the Enterprise OutlineCode1 (Task) field.
    myArray(837) = 188744437  'pjTaskEnterpriseOutlineCode10 -- 188744437 -- Represents the Enterprise OutlineCode10 (Task) field.
    myArray(838) = 188744439  'pjTaskEnterpriseOutlineCode11 -- 188744439 -- Represents the Enterprise OutlineCode11 (Task) field.
    myArray(839) = 188744441  'pjTaskEnterpriseOutlineCode12 -- 188744441 -- Represents the Enterprise OutlineCode12 (Task) field.
    myArray(840) = 188744443  'pjTaskEnterpriseOutlineCode13 -- 188744443 -- Represents the Enterprise OutlineCode13 (Task) field.
    myArray(841) = 188744445  'pjTaskEnterpriseOutlineCode14 -- 188744445 -- Represents the Enterprise OutlineCode14 (Task) field.
    myArray(842) = 188744447  'pjTaskEnterpriseOutlineCode15 -- 188744447 -- Represents the Enterprise OutlineCode15 (Task) field.
    myArray(843) = 188744449  'pjTaskEnterpriseOutlineCode16 -- 188744449 -- Represents the Enterprise OutlineCode16 (Task) field.
    myArray(844) = 188744451  'pjTaskEnterpriseOutlineCode17 -- 188744451 -- Represents the Enterprise OutlineCode17 (Task) field.
    myArray(845) = 188744453  'pjTaskEnterpriseOutlineCode18 -- 188744453 -- Represents the Enterprise OutlineCode18 (Task) field.
    myArray(846) = 188744455  'pjTaskEnterpriseOutlineCode19 -- 188744455 -- Represents the Enterprise OutlineCode19 (Task) field.
    myArray(847) = 188744421  'pjTaskEnterpriseOutlineCode2 -- 188744421 -- Represents the Enterprise OutlineCode2 (Task) field.
    myArray(848) = 188744457  'pjTaskEnterpriseOutlineCode20 -- 188744457 -- Represents the Enterprise OutlineCode20 (Task) field.
    myArray(849) = 188744459  'pjTaskEnterpriseOutlineCode21 -- 188744459 -- Represents the Enterprise OutlineCode21 (Task) field.
    myArray(850) = 188744461  'pjTaskEnterpriseOutlineCode22 -- 188744461 -- Represents the Enterprise OutlineCode22 (Task) field.
    myArray(851) = 188744463  'pjTaskEnterpriseOutlineCode23 -- 188744463 -- Represents the Enterprise OutlineCode23 (Task) field.
    myArray(852) = 188744465  'pjTaskEnterpriseOutlineCode24 -- 188744465 -- Represents the Enterprise OutlineCode24 (Task) field.
    myArray(853) = 188744467  'pjTaskEnterpriseOutlineCode25 -- 188744467 -- Represents the Enterprise OutlineCode25 (Task) field.
    myArray(854) = 188744469  'pjTaskEnterpriseOutlineCode26 -- 188744469 -- Represents the Enterprise OutlineCode26 (Task) field.
    myArray(855) = 188744471  'pjTaskEnterpriseOutlineCode27 -- 188744471 -- Represents the Enterprise OutlineCode27 (Task) field.
    myArray(856) = 188744473  'pjTaskEnterpriseOutlineCode28 -- 188744473 -- Represents the Enterprise OutlineCode28 (Task) field.
    myArray(857) = 188744475  'pjTaskEnterpriseOutlineCode29 -- 188744475 -- Represents the Enterprise OutlineCode29 (Task) field.
    myArray(858) = 188744423  'pjTaskEnterpriseOutlineCode3 -- 188744423 -- Represents the Enterprise OutlineCode3 (Task) field.
    myArray(859) = 188744477  'pjTaskEnterpriseOutlineCode30 -- 188744477 -- Represents the Enterprise OutlineCode30 (Task) field.
    myArray(860) = 188744425  'pjTaskEnterpriseOutlineCode4 -- 188744425 -- Represents the Enterprise OutlineCode4 (Task) field.
    myArray(861) = 188744427  'pjTaskEnterpriseOutlineCode5 -- 188744427 -- Represents the Enterprise OutlineCode5 (Task) field.
    myArray(862) = 188744429  'pjTaskEnterpriseOutlineCode6 -- 188744429 -- Represents the Enterprise OutlineCode6 (Task) field.
    myArray(863) = 188744431  'pjTaskEnterpriseOutlineCode7 -- 188744431 -- Represents the Enterprise OutlineCode7 (Task) field.
    myArray(864) = 188744433  'pjTaskEnterpriseOutlineCode8 -- 188744433 -- Represents the Enterprise OutlineCode8 (Task) field.
    myArray(865) = 188744435  'pjTaskEnterpriseOutlineCode9 -- 188744435 -- Represents the Enterprise OutlineCode9 (Task) field.
    myArray(866) = 188744529  'pjTaskEnterpriseProjectCost1 -- 188744529 -- Represents the Enterprise Project Cost1 (Task) field.
    myArray(867) = 188744538  'pjTaskEnterpriseProjectCost10 -- 188744538 -- Represents the Enterprise Project Cost10 (Task) field.
    myArray(868) = 188744530  'pjTaskEnterpriseProjectCost2 -- 188744530 -- Represents the Enterprise Project Cost2 (Task) field.
    myArray(869) = 188744531  'pjTaskEnterpriseProjectCost3 -- 188744531 -- Represents the Enterprise Project Cost3 (Task) field.
    myArray(870) = 188744532  'pjTaskEnterpriseProjectCost4 -- 188744532 -- Represents the Enterprise Project Cost4 (Task) field.
    myArray(871) = 188744533  'pjTaskEnterpriseProjectCost5 -- 188744533 -- Represents the Enterprise Project Cost5 (Task) field.
    myArray(872) = 188744534  'pjTaskEnterpriseProjectCost6 -- 188744534 -- Represents the Enterprise Project Cost6 (Task) field.
    myArray(873) = 188744535  'pjTaskEnterpriseProjectCost7 -- 188744535 -- Represents the Enterprise Project Cost7 (Task) field.
    myArray(874) = 188744536  'pjTaskEnterpriseProjectCost8 -- 188744536 -- Represents the Enterprise Project Cost8 (Task) field.
    myArray(875) = 188744537  'pjTaskEnterpriseProjectCost9 -- 188744537 -- Represents the Enterprise Project Cost9 (Task) field.
    myArray(876) = 188744539  'pjTaskEnterpriseProjectDate1 -- 188744539 -- Represents the Enterprise Project Date1 (Task) field.
    myArray(877) = 188744548  'pjTaskEnterpriseProjectDate10 -- 188744548 -- Represents the Enterprise Project Date10 (Task) field.
    myArray(878) = 188744549  'pjTaskEnterpriseProjectDate11 -- 188744549 -- Represents the Enterprise Project Date11 (Task) field.
    myArray(879) = 188744550  'pjTaskEnterpriseProjectDate12 -- 188744550 -- Represents the Enterprise Project Date12 (Task) field.
    myArray(880) = 188744551  'pjTaskEnterpriseProjectDate13 -- 188744551 -- Represents the Enterprise Project Date13 (Task) field.
    myArray(881) = 188744552  'pjTaskEnterpriseProjectDate14 -- 188744552 -- Represents the Enterprise Project Date14 (Task) field.
    myArray(882) = 188744553  'pjTaskEnterpriseProjectDate15 -- 188744553 -- Represents the Enterprise Project Date15 (Task) field.
    myArray(883) = 188744554  'pjTaskEnterpriseProjectDate16 -- 188744554 -- Represents the Enterprise Project Date16 (Task) field.
    myArray(884) = 188744555  'pjTaskEnterpriseProjectDate17 -- 188744555 -- Represents the Enterprise Project Date17 (Task) field.
    myArray(885) = 188744556  'pjTaskEnterpriseProjectDate18 -- 188744556 -- Represents the Enterprise Project Date18 (Task) field.
    myArray(886) = 188744557  'pjTaskEnterpriseProjectDate19 -- 188744557 -- Represents the Enterprise Project Date19 (Task) field.
    myArray(887) = 188744540  'pjTaskEnterpriseProjectDate2 -- 188744540 -- Represents the Enterprise Project Date2 (Task) field.
    myArray(888) = 188744558  'pjTaskEnterpriseProjectDate20 -- 188744558 -- Represents the Enterprise Project Date20 (Task) field.
    myArray(889) = 188744559  'pjTaskEnterpriseProjectDate21 -- 188744559 -- Represents the Enterprise Project Date21 (Task) field.
    myArray(890) = 188744560  'pjTaskEnterpriseProjectDate22 -- 188744560 -- Represents the Enterprise Project Date22 (Task) field.
    myArray(891) = 188744561  'pjTaskEnterpriseProjectDate23 -- 188744561 -- Represents the Enterprise Project Date23 (Task) field.
    myArray(892) = 188744562  'pjTaskEnterpriseProjectDate24 -- 188744562 -- Represents the Enterprise Project Date24 (Task) field.
    myArray(893) = 188744563  'pjTaskEnterpriseProjectDate25 -- 188744563 -- Represents the Enterprise Project Date25 (Task) field.
    myArray(894) = 188744564  'pjTaskEnterpriseProjectDate26 -- 188744564 -- Represents the Enterprise Project Date26 (Task) field.
    myArray(895) = 188744565  'pjTaskEnterpriseProjectDate27 -- 188744565 -- Represents the Enterprise Project Date27 (Task) field.
    myArray(896) = 188744566  'pjTaskEnterpriseProjectDate28 -- 188744566 -- Represents the Enterprise Project Date28 (Task) field.
    myArray(897) = 188744567  'pjTaskEnterpriseProjectDate29 -- 188744567 -- Represents the Enterprise Project Date29 (Task) field.
    myArray(898) = 188744541  'pjTaskEnterpriseProjectDate3 -- 188744541 -- Represents the Enterprise Project Date3 (Task) field.
    myArray(899) = 188744568  'pjTaskEnterpriseProjectDate30 -- 188744568 -- Represents the Enterprise Project Date30 (Task) field.
    myArray(900) = 188744542  'pjTaskEnterpriseProjectDate4 -- 188744542 -- Represents the Enterprise Project Date4 (Task) field.
    myArray(901) = 188744543  'pjTaskEnterpriseProjectDate5 -- 188744543 -- Represents the Enterprise Project Date5 (Task) field.
    myArray(902) = 188744544  'pjTaskEnterpriseProjectDate6 -- 188744544 -- Represents the Enterprise Project Date6 (Task) field.
    myArray(903) = 188744545  'pjTaskEnterpriseProjectDate7 -- 188744545 -- Represents the Enterprise Project Date7 (Task) field.
    myArray(904) = 188744546  'pjTaskEnterpriseProjectDate8 -- 188744546 -- Represents the Enterprise Project Date8 (Task) field.
    myArray(905) = 188744547  'pjTaskEnterpriseProjectDate9 -- 188744547 -- Represents the Enterprise Project Date9 (Task) field.
    myArray(906) = 188744569  'pjTaskEnterpriseProjectDuration1 -- 188744569 -- Represents the Enterprise Project Duration1 (Task) field.
    myArray(907) = 188744578  'pjTaskEnterpriseProjectDuration10 -- 188744578 -- Represents the Enterprise Project Duration10 (Task) field.
    myArray(908) = 188744570  'pjTaskEnterpriseProjectDuration2 -- 188744570 -- Represents the Enterprise Project Duration2 (Task) field.
    myArray(909) = 188744571  'pjTaskEnterpriseProjectDuration3 -- 188744571 -- Represents the Enterprise Project Duration3 (Task) field.
    myArray(910) = 188744572  'pjTaskEnterpriseProjectDuration4 -- 188744572 -- Represents the Enterprise Project Duration4 (Task) field.
    myArray(911) = 188744573  'pjTaskEnterpriseProjectDuration5 -- 188744573 -- Represents the Enterprise Project Duration5 (Task) field.
    myArray(912) = 188744574  'pjTaskEnterpriseProjectDuration6 -- 188744574 -- Represents the Enterprise Project Duration6 (Task) field.
    myArray(913) = 188744575  'pjTaskEnterpriseProjectDuration7 -- 188744575 -- Represents the Enterprise Project Duration7 (Task) field.
    myArray(914) = 188744576  'pjTaskEnterpriseProjectDuration8 -- 188744576 -- Represents the Enterprise Project Duration8 (Task) field.
    myArray(915) = 188744577  'pjTaskEnterpriseProjectDuration9 -- 188744577 -- Represents the Enterprise Project Duration9 (Task) field.
    myArray(916) = 188744649  'pjTaskEnterpriseProjectFlag1 -- 188744649 -- Represents the Enterprise Project Flag1 (Task) field.
    myArray(917) = 188744658  'pjTaskEnterpriseProjectFlag10 -- 188744658 -- Represents the Enterprise Project Flag10 (Task) field.
    myArray(918) = 188744659  'pjTaskEnterpriseProjectFlag11 -- 188744659 -- Represents the Enterprise Project Flag11 (Task) field.
    myArray(919) = 188744660  'pjTaskEnterpriseProjectFlag12 -- 188744660 -- Represents the Enterprise Project Flag12 (Task) field.
    myArray(920) = 188744661  'pjTaskEnterpriseProjectFlag13 -- 188744661 -- Represents the Enterprise Project Flag13 (Task) field.
    myArray(921) = 188744662  'pjTaskEnterpriseProjectFlag14 -- 188744662 -- Represents the Enterprise Project Flag14 (Task) field.
    myArray(922) = 188744663  'pjTaskEnterpriseProjectFlag15 -- 188744663 -- Represents the Enterprise Project Flag15 (Task) field.
    myArray(923) = 188744664  'pjTaskEnterpriseProjectFlag16 -- 188744664 -- Represents the Enterprise Project Flag16 (Task) field.
    myArray(924) = 188744665  'pjTaskEnterpriseProjectFlag17 -- 188744665 -- Represents the Enterprise Project Flag17 (Task) field.
    myArray(925) = 188744666  'pjTaskEnterpriseProjectFlag18 -- 188744666 -- Represents the Enterprise Project Flag18 (Task) field.
    myArray(926) = 188744667  'pjTaskEnterpriseProjectFlag19 -- 188744667 -- Represents the Enterprise Project Flag19 (Task) field.
    myArray(927) = 188744650  'pjTaskEnterpriseProjectFlag2 -- 188744650 -- Represents the Enterprise Project Flag2 (Task) field.
    myArray(928) = 188744668  'pjTaskEnterpriseProjectFlag20 -- 188744668 -- Represents the Enterprise Project Flag20 (Task) field.
    myArray(929) = 188744651  'pjTaskEnterpriseProjectFlag3 -- 188744651 -- Represents the Enterprise Project Flag3 (Task) field.
    myArray(930) = 188744652  'pjTaskEnterpriseProjectFlag4 -- 188744652 -- Represents the Enterprise Project Flag4 (Task) field.
    myArray(931) = 188744653  'pjTaskEnterpriseProjectFlag5 -- 188744653 -- Represents the Enterprise Project Flag5 (Task) field.
    myArray(932) = 188744654  'pjTaskEnterpriseProjectFlag6 -- 188744654 -- Represents the Enterprise Project Flag6 (Task) field.
    myArray(933) = 188744655  'pjTaskEnterpriseProjectFlag7 -- 188744655 -- Represents the Enterprise Project Flag7 (Task) field.
    myArray(934) = 188744656  'pjTaskEnterpriseProjectFlag8 -- 188744656 -- Represents the Enterprise Project Flag8 (Task) field.
    myArray(935) = 188744657  'pjTaskEnterpriseProjectFlag9 -- 188744657 -- Represents the Enterprise Project Flag9 (Task) field.
    myArray(936) = 188744689  'pjTaskEnterpriseProjectNumber1 -- 188744689 -- Represents the Enterprise Project Number1 (Task) field.
    myArray(937) = 188744698  'pjTaskEnterpriseProjectNumber10 -- 188744698 -- Represents the Enterprise Project Number10 (Task) field.
    myArray(938) = 188744699  'pjTaskEnterpriseProjectNumber11 -- 188744699 -- Represents the Enterprise Project Number11 (Task) field.
    myArray(939) = 188744700  'pjTaskEnterpriseProjectNumber12 -- 188744700 -- Represents the Enterprise Project Number12 (Task) field.
    myArray(940) = 188744701  'pjTaskEnterpriseProjectNumber13 -- 188744701 -- Represents the Enterprise Project Number13 (Task) field.
    myArray(941) = 188744702  'pjTaskEnterpriseProjectNumber14 -- 188744702 -- Represents the Enterprise Project Number14 (Task) field.
    myArray(942) = 188744703  'pjTaskEnterpriseProjectNumber15 -- 188744703 -- Represents the Enterprise Project Number15 (Task) field.
    myArray(943) = 188744704  'pjTaskEnterpriseProjectNumber16 -- 188744704 -- Represents the Enterprise Project Number16 (Task) field.
    myArray(944) = 188744705  'pjTaskEnterpriseProjectNumber17 -- 188744705 -- Represents the Enterprise Project Number17 (Task) field.
    myArray(945) = 188744706  'pjTaskEnterpriseProjectNumber18 -- 188744706 -- Represents the Enterprise Project Number18 (Task) field.
    myArray(946) = 188744707  'pjTaskEnterpriseProjectNumber19 -- 188744707 -- Represents the Enterprise Project Number19 (Task) field.
    myArray(947) = 188744690  'pjTaskEnterpriseProjectNumber2 -- 188744690 -- Represents the Enterprise Project Number2 (Task) field.
    myArray(948) = 188744708  'pjTaskEnterpriseProjectNumber20 -- 188744708 -- Represents the Enterprise Project Number20 (Task) field.
    myArray(949) = 188744709  'pjTaskEnterpriseProjectNumber21 -- 188744709 -- Represents the Enterprise Project Number21 (Task) field.
    myArray(950) = 188744710  'pjTaskEnterpriseProjectNumber22 -- 188744710 -- Represents the Enterprise Project Number22 (Task) field.
    myArray(951) = 188744711  'pjTaskEnterpriseProjectNumber23 -- 188744711 -- Represents the Enterprise Project Number23 (Task) field.
    myArray(952) = 188744712  'pjTaskEnterpriseProjectNumber24 -- 188744712 -- Represents the Enterprise Project Number24 (Task) field.
    myArray(953) = 188744713  'pjTaskEnterpriseProjectNumber25 -- 188744713 -- Represents the Enterprise Project Number25 (Task) field.
    myArray(954) = 188744714  'pjTaskEnterpriseProjectNumber26 -- 188744714 -- Represents the Enterprise Project Number26 (Task) field.
    myArray(955) = 188744715  'pjTaskEnterpriseProjectNumber27 -- 188744715 -- Represents the Enterprise Project Number27 (Task) field.
    myArray(956) = 188744716  'pjTaskEnterpriseProjectNumber28 -- 188744716 -- Represents the Enterprise Project Number28 (Task) field.
    myArray(957) = 188744717  'pjTaskEnterpriseProjectNumber29 -- 188744717 -- Represents the Enterprise Project Number29 (Task) field.
    myArray(958) = 188744691  'pjTaskEnterpriseProjectNumber3 -- 188744691 -- Represents the Enterprise Project Number3 (Task) field.
    myArray(959) = 188744718  'pjTaskEnterpriseProjectNumber30 -- 188744718 -- Represents the Enterprise Project Number30 (Task) field.
    myArray(960) = 188744719  'pjTaskEnterpriseProjectNumber31 -- 188744719 -- Represents the Enterprise Project Number31 (Task) field.
    myArray(961) = 188744720  'pjTaskEnterpriseProjectNumber32 -- 188744720 -- Represents the Enterprise Project Number32 (Task) field.
    myArray(962) = 188744721  'pjTaskEnterpriseProjectNumber33 -- 188744721 -- Represents the Enterprise Project Number33 (Task) field.
    myArray(963) = 188744722  'pjTaskEnterpriseProjectNumber34 -- 188744722 -- Represents the Enterprise Project Number34 (Task) field.
    myArray(964) = 188744723  'pjTaskEnterpriseProjectNumber35 -- 188744723 -- Represents the Enterprise Project Number35 (Task) field.
    myArray(965) = 188744724  'pjTaskEnterpriseProjectNumber36 -- 188744724 -- Represents the Enterprise Project Number36 (Task) field.
    myArray(966) = 188744725  'pjTaskEnterpriseProjectNumber37 -- 188744725 -- Represents the Enterprise Project Number37 (Task) field.
    myArray(967) = 188744726  'pjTaskEnterpriseProjectNumber38 -- 188744726 -- Represents the Enterprise Project Number38 (Task) field.
    myArray(968) = 188744727  'pjTaskEnterpriseProjectNumber39 -- 188744727 -- Represents the Enterprise Project Number39 (Task) field.
    myArray(969) = 188744692  'pjTaskEnterpriseProjectNumber4 -- 188744692 -- Represents the Enterprise Project Number4(Task) field.
    myArray(970) = 188744728  'pjTaskEnterpriseProjectNumber40 -- 188744728 -- Represents the Enterprise Project Number40 (Task) field.
    myArray(971) = 188744693  'pjTaskEnterpriseProjectNumber5 -- 188744693 -- Represents the Enterprise Project Number5 (Task) field.
    myArray(972) = 188744694  'pjTaskEnterpriseProjectNumber6 -- 188744694 -- Represents the Enterprise Project Number6 (Task) field.
    myArray(973) = 188744695  'pjTaskEnterpriseProjectNumber7 -- 188744695 -- Represents the Enterprise Project Number7 (Task) field.
    myArray(974) = 188744696  'pjTaskEnterpriseProjectNumber8 -- 188744696 -- Represents the Enterprise Project Number8 (Task) field.
    myArray(975) = 188744697  'pjTaskEnterpriseProjectNumber9 -- 188744697 -- Represents the Enterprise Project Number9 (Task) field.
    myArray(976) = 188744589  'pjTaskEnterpriseProjectOutlineCode1 -- 188744589 -- Represents the Enterprise Project OutlineCode1 (Task) field.
    myArray(977) = 188744598  'pjTaskEnterpriseProjectOutlineCode10 -- 188744598 -- Represents the Enterprise Project OutlineCode10 (Task) field.
    myArray(978) = 188744599  'pjTaskEnterpriseProjectOutlineCode11 -- 188744599 -- Represents the Enterprise Project OutlineCode11 (Task) field.
    myArray(979) = 188744600  'pjTaskEnterpriseProjectOutlineCode12 -- 188744600 -- Represents the Enterprise Project OutlineCode12 (Task) field.
    myArray(980) = 188744601  'pjTaskEnterpriseProjectOutlineCode13 -- 188744601 -- Represents the Enterprise Project OutlineCode13 (Task) field.
    myArray(981) = 188744602  'pjTaskEnterpriseProjectOutlineCode14 -- 188744602 -- Represents the Enterprise Project OutlineCode14 (Task) field.
    myArray(982) = 188744603  'pjTaskEnterpriseProjectOutlineCode15 -- 188744603 -- Represents the Enterprise Project OutlineCode15 (Task) field.
    myArray(983) = 188744604  'pjTaskEnterpriseProjectOutlineCode16 -- 188744604 -- Represents the Enterprise Project OutlineCode16 (Task) field.
    myArray(984) = 188744605  'pjTaskEnterpriseProjectOutlineCode17 -- 188744605 -- Represents the Enterprise Project OutlineCode17 (Task) field.
    myArray(985) = 188744606  'pjTaskEnterpriseProjectOutlineCode18 -- 188744606 -- Represents the Enterprise Project OutlineCode18 (Task) field.
    myArray(986) = 188744607  'pjTaskEnterpriseProjectOutlineCode19 -- 188744607 -- Represents the Enterprise Project OutlineCode19 (Task) field.
    myArray(987) = 188744590  'pjTaskEnterpriseProjectOutlineCode2 -- 188744590 -- Represents the Enterprise Project OutlineCode2 (Task) field.
    myArray(988) = 188744608  'pjTaskEnterpriseProjectOutlineCode20 -- 188744608 -- Represents the Enterprise Project OutlineCode2 (Task) field.
    myArray(989) = 188744609  'pjTaskEnterpriseProjectOutlineCode21 -- 188744609 -- Represents the Enterprise Project OutlineCode21 (Task) field.
    myArray(990) = 188744610  'pjTaskEnterpriseProjectOutlineCode22 -- 188744610 -- Represents the Enterprise Project OutlineCode22 (Task) field.
    myArray(991) = 188744611  'pjTaskEnterpriseProjectOutlineCode23 -- 188744611 -- Represents the Enterprise Project OutlineCode23 (Task) field.
    myArray(992) = 188744612  'pjTaskEnterpriseProjectOutlineCode24 -- 188744612 -- Represents the Enterprise Project OutlineCode24 (Task) field.
    myArray(993) = 188744613  'pjTaskEnterpriseProjectOutlineCode25 -- 188744613 -- Represents the Enterprise Project OutlineCode25 (Task) field.
    myArray(994) = 188744614  'pjTaskEnterpriseProjectOutlineCode26 -- 188744614 -- Represents the Enterprise Project OutlineCode26 (Task) field.
    myArray(995) = 188744615  'pjTaskEnterpriseProjectOutlineCode27 -- 188744615 -- Represents the Enterprise Project OutlineCode27 (Task) field.
    myArray(996) = 188744616  'pjTaskEnterpriseProjectOutlineCode28 -- 188744616 -- Represents the Enterprise Project OutlineCode28 (Task) field.
    myArray(997) = 188744617  'pjTaskEnterpriseProjectOutlineCode29 -- 188744617 -- Represents the Enterprise Project OutlineCode29 (Task) field.
    myArray(998) = 188744591  'pjTaskEnterpriseProjectOutlineCode3 -- 188744591 -- Represents the Enterprise Project OutlineCode3 (Task) field.
    myArray(999) = 188744618  'pjTaskEnterpriseProjectOutlineCode30 -- 188744618 -- Represents the Enterprise Project OutlineCode30 (Task) field.
    myArray(1000) = 188744592 'pjTaskEnterpriseProjectOutlineCode4 -- 188744592 -- Represents the Enterprise Project OutlineCode4 (Task) field.
    myArray(1001) = 188744593 'pjTaskEnterpriseProjectOutlineCode5 -- 188744593 -- Represents the Enterprise Project OutlineCode5 (Task) field.
    myArray(1002) = 188744594 'pjTaskEnterpriseProjectOutlineCode6 -- 188744594 -- Represents the Enterprise Project OutlineCode6 (Task) field.
    myArray(1003) = 188744595 'pjTaskEnterpriseProjectOutlineCode7 -- 188744595 -- Represents the Enterprise Project OutlineCode7 (Task) field.
    myArray(1004) = 188744596 'pjTaskEnterpriseProjectOutlineCode8 -- 188744596 -- Represents the Enterprise Project OutlineCode8 (Task) field.
    myArray(1005) = 188744597 'pjTaskEnterpriseProjectOutlineCode9 -- 188744597 -- Represents the Enterprise Project OutlineCode9 (Task) field.
    myArray(1006) = 188744729 'pjTaskEnterpriseProjectText1 -- 188744729 -- Represents the Enterprise Project Text1 (Task) field.
    myArray(1007) = 188744738 'pjTaskEnterpriseProjectText10 -- 188744738 -- Represents the Enterprise Project Text10 (Task) field.
    myArray(1008) = 188744739 'pjTaskEnterpriseProjectText11 -- 188744739 -- Represents the Enterprise Project Text11 (Task) field.
    myArray(1009) = 188744740 'pjTaskEnterpriseProjectText12 -- 188744740 -- Represents the Enterprise Project Text12 (Task) field.
    myArray(1010) = 188744741 'pjTaskEnterpriseProjectText13 -- 188744741 -- Represents the Enterprise Project Text13 (Task) field.
    myArray(1011) = 188744742 'pjTaskEnterpriseProjectText14 -- 188744742 -- Represents the Enterprise Project Text14 (Task) field.
    myArray(1012) = 188744743 'pjTaskEnterpriseProjectText15 -- 188744743 -- Represents the Enterprise Project Text15 (Task) field.
    myArray(1013) = 188744744 'pjTaskEnterpriseProjectText16 -- 188744744 -- Represents the Enterprise Project Text16 (Task) field.
    myArray(1014) = 188744745 'pjTaskEnterpriseProjectText17 -- 188744745 -- Represents the Enterprise Project Text17 (Task) field.
    myArray(1015) = 188744746 'pjTaskEnterpriseProjectText18 -- 188744746 -- Represents the Enterprise Project Text18 (Task) field.
    myArray(1016) = 188744747 'pjTaskEnterpriseProjectText19 -- 188744747 -- Represents the Enterprise Project Text19 (Task) field.
    myArray(1017) = 188744730 'pjTaskEnterpriseProjectText2 -- 188744730 -- Represents the Enterprise Project Text2 (Task) field.
    myArray(1018) = 188744748 'pjTaskEnterpriseProjectText20 -- 188744748 -- Represents the Enterprise Project Text20 (Task) field.
    myArray(1019) = 188744749 'pjTaskEnterpriseProjectText21 -- 188744749 -- Represents the Enterprise Project Text21 (Task) field.
    myArray(1020) = 188744750 'pjTaskEnterpriseProjectText22 -- 188744750 -- Represents the Enterprise Project Text22 (Task) field.
    myArray(1021) = 188744751 'pjTaskEnterpriseProjectText23 -- 188744751 -- Represents the Enterprise Project Text23 (Task) field.
    myArray(1022) = 188744752 'pjTaskEnterpriseProjectText24 -- 188744752 -- Represents the Enterprise Project Text24 (Task) field.
    myArray(1023) = 188744753 'pjTaskEnterpriseProjectText25 -- 188744753 -- Represents the Enterprise Project Text25 (Task) field.
    myArray(1024) = 188744754 'pjTaskEnterpriseProjectText26 -- 188744754 -- Represents the Enterprise Project Text26 (Task) field.
    myArray(1025) = 188744755 'pjTaskEnterpriseProjectText27 -- 188744755 -- Represents the Enterprise Project Text27 (Task) field.
    myArray(1026) = 188744756 'pjTaskEnterpriseProjectText28 -- 188744756 -- Represents the Enterprise Project Text28 (Task) field.
    myArray(1027) = 188744757 'pjTaskEnterpriseProjectText29 -- 188744757 -- Represents the Enterprise Project Text29 (Task) field.
    myArray(1028) = 188744731 'pjTaskEnterpriseProjectText3 -- 188744731 -- Represents the Enterprise Project Text3 (Task) field.
    myArray(1029) = 188744758 'pjTaskEnterpriseProjectText30 -- 188744758 -- Represents the Enterprise Project Text30 (Task) field.
    myArray(1030) = 188744759 'pjTaskEnterpriseProjectText31 -- 188744759 -- Represents the Enterprise Project Text31 (Task) field.
    myArray(1031) = 188744760 'pjTaskEnterpriseProjectText32 -- 188744760 -- Represents the Enterprise Project Text32 (Task) field.
    myArray(1032) = 188744761 'pjTaskEnterpriseProjectText33 -- 188744761 -- Represents the Enterprise Project Text33 (Task) field.
    myArray(1033) = 188744762 'pjTaskEnterpriseProjectText34 -- 188744762 -- Represents the Enterprise Project Text34 (Task) field.
    myArray(1034) = 188744763 'pjTaskEnterpriseProjectText35 -- 188744763 -- Represents the Enterprise Project Text35 (Task) field.
    myArray(1035) = 188744764 'pjTaskEnterpriseProjectText36 -- 188744764 -- Represents the Enterprise Project Text36 (Task) field.
    myArray(1036) = 188744765 'pjTaskEnterpriseProjectText37 -- 188744765 -- Represents the Enterprise Project Text37 (Task) field.
    myArray(1037) = 188744766 'pjTaskEnterpriseProjectText38 -- 188744766 -- Represents the Enterprise Project Text38 (Task) field.
    myArray(1038) = 188744767 'pjTaskEnterpriseProjectText39 -- 188744767 -- Represents the Enterprise Project Text39 (Task) field.
    myArray(1039) = 188744732 'pjTaskEnterpriseProjectText4 -- 188744732 -- Represents the Enterprise Project Text4 (Task) field.
    myArray(1040) = 188744768 'pjTaskEnterpriseProjectText40 -- 188744768 -- Represents the Enterprise Project Text40 (Task) field.
    myArray(1041) = 188744733 'pjTaskEnterpriseProjectText5 -- 188744733 -- Represents the Enterprise Project Text5 (Task) field.
    myArray(1042) = 188744734 'pjTaskEnterpriseProjectText6 -- 188744734 -- Represents the Enterprise Project Text6 (Task) field.
    myArray(1043) = 188744735 'pjTaskEnterpriseProjectText7 -- 188744735 -- Represents the Enterprise Project Text7 (Task) field.
    myArray(1044) = 188744736 'pjTaskEnterpriseProjectText8 -- 188744736 -- Represents the Enterprise Project Text8 (Task) field.
    myArray(1045) = 188744737 'pjTaskEnterpriseProjectText9 -- 188744737 -- Represents the Enterprise Project Text9(Task) field.
    myArray(1046) = 188744479 'pjTaskEnterpriseText1 -- 188744479 -- Represents the Enterprise Text1 (Task) field.
    myArray(1047) = 188744488 'pjTaskEnterpriseText10 -- 188744488 -- Represents the Enterprise Text10 (Task) field.
    myArray(1048) = 188744489 'pjTaskEnterpriseText11 -- 188744489 -- Represents the Enterprise Text11 (Task) field.
    myArray(1049) = 188744490 'pjTaskEnterpriseText12 -- 188744490 -- Represents the Enterprise Text12 (Task) field.
    myArray(1050) = 188744491 'pjTaskEnterpriseText13 -- 188744491 -- Represents the Enterprise Text13 (Task) field.
    myArray(1051) = 188744492 'pjTaskEnterpriseText14 -- 188744492 -- Represents the Enterprise Text14 (Task) field.
    myArray(1052) = 188744493 'pjTaskEnterpriseText15 -- 188744493 -- Represents the Enterprise Text15 (Task) field.
    myArray(1053) = 188744494 'pjTaskEnterpriseText16 -- 188744494 -- Represents the Enterprise Text16 (Task) field.
    myArray(1054) = 188744495 'pjTaskEnterpriseText17 -- 188744495 -- Represents the Enterprise Text17 (Task) field.
    myArray(1055) = 188744496 'pjTaskEnterpriseText18 -- 188744496 -- Represents the Enterprise Text18 (Task) field.
    myArray(1056) = 188744497 'pjTaskEnterpriseText19 -- 188744497 -- Represents the Enterprise Text19 (Task) field.
    myArray(1057) = 188744480 'pjTaskEnterpriseText2 -- 188744480 -- Represents the Enterprise Text2 (Task) field.
    myArray(1058) = 188744498 'pjTaskEnterpriseText20 -- 188744498 -- Represents the Enterprise Text20 (Task) field.
    myArray(1059) = 188744499 'pjTaskEnterpriseText21 -- 188744499 -- Represents the Enterprise Text21 (Task) field.
    myArray(1060) = 188744500 'pjTaskEnterpriseText22 -- 188744500 -- Represents the Enterprise Text22 (Task) field.
    myArray(1061) = 188744501 'pjTaskEnterpriseText23 -- 188744501 -- Represents the Enterprise Text23 (Task) field.
    myArray(1062) = 188744502 'pjTaskEnterpriseText24 -- 188744502 -- Represents the Enterprise Text24 (Task) field.
    myArray(1063) = 188744503 'pjTaskEnterpriseText25 -- 188744503 -- Represents the Enterprise Text25 (Task) field.
    myArray(1064) = 188744504 'pjTaskEnterpriseText26 -- 188744504 -- Represents the Enterprise Text26 (Task) field.
    myArray(1065) = 188744505 'pjTaskEnterpriseText27 -- 188744505 -- Represents the Enterprise Text27 (Task) field.
    myArray(1066) = 188744506 'pjTaskEnterpriseText28 -- 188744506 -- Represents the Enterprise Text28 (Task) field.
    myArray(1067) = 188744507 'pjTaskEnterpriseText29 -- 188744507 -- Represents the Enterprise Text29 (Task) field.
    myArray(1068) = 188744481 'pjTaskEnterpriseText3 -- 188744481 -- Represents the Enterprise Text3 (Task) field.
    myArray(1069) = 188744508 'pjTaskEnterpriseText30 -- 188744508 -- Represents the Enterprise Text30 (Task) field.
    myArray(1070) = 188744509 'pjTaskEnterpriseText31 -- 188744509 -- Represents the Enterprise Text31 (Task) field.
    myArray(1071) = 188744510 'pjTaskEnterpriseText32 -- 188744510 -- Represents the Enterprise Text32 (Task) field.
    myArray(1072) = 188744511 'pjTaskEnterpriseText33 -- 188744511 -- Represents the Enterprise Text33 (Task) field.
    myArray(1073) = 188744512 'pjTaskEnterpriseText34 -- 188744512 -- Represents the Enterprise Text34 (Task) field.
    myArray(1074) = 188744513 'pjTaskEnterpriseText35 -- 188744513 -- Represents the Enterprise Text35 (Task) field.
    myArray(1075) = 188744514 'pjTaskEnterpriseText36 -- 188744514 -- Represents the Enterprise Text36 (Task) field.
    myArray(1076) = 188744515 'pjTaskEnterpriseText37 -- 188744515 -- Represents the Enterprise Text37 (Task) field.
    myArray(1077) = 188744516 'pjTaskEnterpriseText38 -- 188744516 -- Represents the Enterprise Text38 (Task) field.
    myArray(1078) = 188744517 'pjTaskEnterpriseText39 -- 188744517 -- Represents the Enterprise Text39 (Task) field.
    myArray(1079) = 188744482 'pjTaskEnterpriseText4 -- 188744482 -- Represents the Enterprise Text4 (Task) field.
    myArray(1080) = 188744518 'pjTaskEnterpriseText40 -- 188744518 -- Represents the Enterprise Text40 (Task) field.
    myArray(1081) = 188744483 'pjTaskEnterpriseText5 -- 188744483 -- Represents the Enterprise Text5 (Task) field.
    myArray(1082) = 188744484 'pjTaskEnterpriseText6 -- 188744484 -- Represents the Enterprise Text6 (Task) field.
    myArray(1083) = 188744485 'pjTaskEnterpriseText7 -- 188744485 -- Represents the Enterprise Text7 (Task) field.
    myArray(1084) = 188744486 'pjTaskEnterpriseText8 -- 188744486 -- Represents the Enterprise Text8 (Task) field.
    myArray(1085) = 188744487 'pjTaskEnterpriseText9 -- 188744487 -- Represents the Enterprise Text9 (Task) field.
    myArray(1086) = 188744847 'pjTaskErrorMessage -- 188744847 -- Represents the Error Message (Task) field.
    myArray(1087) = 188744076 'pjTaskEstimated -- 188744076 -- Represents the Estimated (Task) field.
    myArray(1088) = 188743912 'pjTaskExternalTask -- 188743912 -- Represents the External Task (Task) field.
    myArray(1089) = 188743716 'pjTaskFinish -- 188743716 -- Represents the Finish (Task) field.
    myArray(1090) = 188743733 'pjTaskFinish1 -- 188743733 -- Represents the Finish1 (Task) field.
    myArray(1091) = 188743971 'pjTaskFinish10 -- 188743971 -- Represents the Finish10 (Task) field.
    myArray(1092) = 188743736 'pjTaskFinish2 -- 188743736 -- Represents the Finish2 (Task) field.
    myArray(1093) = 188743739 'pjTaskFinish3 -- 188743739 -- Represents the Finish3 (Task) field.
    myArray(1094) = 188743742 'pjTaskFinish4 -- 188743742 -- Represents the Finish4 (Task) field.
    myArray(1095) = 188743745 'pjTaskFinish5 -- 188743745 -- Represents the Finish5 (Task) field.
    myArray(1096) = 188743963 'pjTaskFinish6 -- 188743963 -- Represents the Finish6 (Task) field.
    myArray(1097) = 188743965 'pjTaskFinish7 -- 188743965 -- Represents the Finish7 (Task) field.
    myArray(1098) = 188743967 'pjTaskFinish8 -- 188743967 -- Represents the Finish8 (Task) field.
    myArray(1099) = 188743969 'pjTaskFinish9 -- 188743969 -- Represents the Finish9 (Task) field.
    myArray(1100) = 188744119 'pjTaskFinishSlack -- 188744119 -- Represents the Finish Slack (Task) field.
    myArray(1101) = 188744966 'pjTaskFinishText -- 188744966 -- Represents the Task Finish Text (Task) field.
    myArray(1102) = 188743726 'pjTaskFinishVariance -- 188743726 -- Represents the Finish Variance (Task) field.
    myArray(1103) = 188743688 'pjTaskFixedCost -- 188743688 -- Represents the Fixed Cost (Task) field.
    myArray(1104) = 188743880 'pjTaskFixedCostAccrual -- 188743880 -- Represents the Fixed Cost Accrual (Task) field.
    myArray(1105) = 188743714 'pjTaskFixedDuration -- 188743714 -- Represents the Fixed Duration (Task) field.
    myArray(1106) = 188743752 'pjTaskFlag1 -- 188743752 -- Represents the Flag1 (Task) field.
    myArray(1107) = 188743761 'pjTaskFlag10 -- 188743761 -- Represents the Flag10 (Task) field.
    myArray(1108) = 188743972 'pjTaskFlag11 -- 188743972 -- Represents the Flag11 (Task) field.
    myArray(1109) = 188743973 'pjTaskFlag12 -- 188743973 -- Represents the Flag12 (Task) field.
    myArray(1110) = 188743974 'pjTaskFlag13 -- 188743974 -- Represents the Flag13 (Task) field.
    myArray(1111) = 188743975 'pjTaskFlag14 -- 188743975 -- Represents the Flag14 (Task) field.
    myArray(1112) = 188743976 'pjTaskFlag15 -- 188743976 -- Represents the Flag15 (Task) field.
    myArray(1113) = 188743977 'pjTaskFlag16 -- 188743977 -- Represents the Flag16 (Task) field.
    myArray(1114) = 188743978 'pjTaskFlag17 -- 188743978 -- Represents the Flag17 (Task) field.
    myArray(1115) = 188743979 'pjTaskFlag18 -- 188743979 -- Represents the Flag18 (Task) field.
    myArray(1116) = 188743980 'pjTaskFlag19 -- 188743980 -- Represents the Flag19 (Task) field.
    myArray(1117) = 188743753 'pjTaskFlag2 -- 188743753 -- Represents the Flag2 (Task) field.
    myArray(1118) = 188743981 'pjTaskFlag20 -- 188743981 -- Represents the Flag20 (Task) field.
    myArray(1119) = 188743754 'pjTaskFlag3 -- 188743754 -- Represents the Flag3 (Task) field.
    myArray(1120) = 188743755 'pjTaskFlag4 -- 188743755 -- Represents the Flag4 (Task) field.
    myArray(1121) = 188743756 'pjTaskFlag5 -- 188743756 -- Represents the Flag5 (Task) field.
    myArray(1122) = 188743757 'pjTaskFlag6 -- 188743757 -- Represents the Flag6 (Task) field.
    myArray(1123) = 188743758 'pjTaskFlag7 -- 188743758 -- Represents the Flag7 (Task) field.
    myArray(1124) = 188743759 'pjTaskFlag8 -- 188743759 -- Represents the Flag8 (Task) field.
    myArray(1125) = 188743760 'pjTaskFlag9 -- 188743760 -- Represents the Flag9 (Task) field.
    myArray(1126) = 188743701 'pjTaskFreeSlack -- 188743701 -- Represents the Free Slack (Task) field.
    myArray(1127) = 188744126 'pjTaskGroupBySummary -- 188744126 -- Represents the Group By Summary (Task) field.
    myArray(1128) = 188744823 'pjTaskGuid -- 188744823 -- Represents the Task GUID (Task) field.
    myArray(1129) = 188743789 'pjTaskHideBar -- 188743789 -- Represents the Hide Bar (Task) field.
    myArray(1130) = 188743897 'pjTaskHyperlink -- 188743897 -- Represents the Hyperlink (Task) field.
    myArray(1131) = 188743898 'pjTaskHyperlinkAddress -- 188743898 -- Represents the Hyperlink Address (Task) field.
    myArray(1132) = 188743900 'pjTaskHyperlinkHref -- 188743900 -- Represents the Hyperlink HRef (Task) field.
    myArray(1133) = 188744132 'pjTaskHyperlinkScreenTip -- 188744132 -- Represents the Hyperlink ScreenTip (Task) field.
    myArray(1134) = 188743899 'pjTaskHyperlinkSubAddress -- 188743899 -- Represents the Hyperlink Sub Address (Task) field.
    myArray(1135) = 188743703 'pjTaskID -- 188743703 -- Represents the ID (Task) field.
    myArray(1136) = 188744079 'pjTaskIgnoreResourceCalendar -- 188744079 -- Represents the Ignore Resource Calendar (Task) field.
    myArray(1137) = 188745012 'pjTaskIgnoreWarnings -- 188745012 -- Represents the Task Ignore Warnings (Task) field.
    myArray(1138) = 188744040 'pjTaskIndex -- 188744040 -- Represents the Index (Task) field.
    myArray(1139) = 188743885 'pjTaskIndicators -- 188743885 -- Represents the Indicators (Task) field.
    myArray(1140) = 188743904 'pjTaskIsAssignment -- 188743904 -- Represents the Is Assignment (Task) field.
    myArray(1141) = 188744977 'pjTaskIsDurationValid -- 188744977 -- Represents the field that shows whether task duration is valid, when converting from a manual to an automatically scheduled task.
    myArray(1142) = 188744976 'pjTaskIsFinishValid -- 188744976 -- Represents the field that shows whether task finish date is valid, when converting from a manual to an automatically scheduled task.
    myArray(1143) = 188744975 'pjTaskIsStartValid -- 188744975 -- Represents the field that shows whether task start date is valid, when converting from a manual to an automatically scheduled task.
    myArray(1144) = 188744845 'pjTaskIsPublished -- 188744845 -- Represents the Is Published (Task) field.
    myArray(1145) = 188743720 'pjTaskLateFinish -- 188743720 -- Represents the Late Finish (Task) field.
    myArray(1146) = 188743719 'pjTaskLateStart -- 188743719 -- Represents the Late Start (Task) field.
    myArray(1147) = 188743933 'pjTaskLevelAssignments -- 188743933 -- Represents the Level Assignments (Task) field.
    myArray(1148) = 188743932 'pjTaskLevelCanSplit -- 188743932 -- Represents the Level Can Split (Task) field.
    myArray(1149) = 188743700 'pjTaskLevelDelay -- 188743700 -- Represents the Level Delay (Task) field.
    myArray(1150) = 188743778 'pjTaskLinkedFields -- 188743778 -- Represents the Linked Fields (Task) field.
    myArray(1151) = 188744960 'pjTaskManual -- 188744960 -- Represents the Task Manual (TASK) field.
    myArray(1152) = 188743751 'pjTaskMarked -- 188743751 -- Represents the Marked (Task) field.
    myArray(1153) = 188743704 'pjTaskMilestone -- 188743704 -- Represents the Milestone (Task) field.
    myArray(1154) = 188743694 'pjTaskName -- 188743694 -- Represents the Name (Task) field.
    myArray(1155) = 188743695 'pjTaskNotes -- 188743695 -- Represents the Notes (Task) field.
    myArray(1156) = 188743767 'pjTaskNumber1 -- 188743767 -- Represents the Number1 (Task) field.
    myArray(1157) = 188743986 'pjTaskNumber10 -- 188743986 -- Represents the Number10 (Task) field.
    myArray(1158) = 188743987 'pjTaskNumber11 -- 188743987 -- Represents the Number11 (Task) field.
    myArray(1159) = 188743988 'pjTaskNumber12 -- 188743988 -- Represents the Number12 (Task) field.
    myArray(1160) = 188743989 'pjTaskNumber13 -- 188743989 -- Represents the Number13 (Task) field.
    myArray(1161) = 188743990 'pjTaskNumber14 -- 188743990 -- Represents the Number14 (Task) field.
    myArray(1162) = 188743991 'pjTaskNumber15 -- 188743991 -- Represents the Number15 (Task) field..
    myArray(1163) = 188743992 'pjTaskNumber16 -- 188743992 -- Represents the Number16 (Task) field.
    myArray(1164) = 188743993 'pjTaskNumber17 -- 188743993 -- Represents the Number17 (Task) field.
    myArray(1165) = 188743994 'pjTaskNumber18 -- 188743994 -- Represents the Number18 (Task) field.
    myArray(1166) = 188743995 'pjTaskNumber19 -- 188743995 -- Represents the Number19 (Task) field.
    myArray(1167) = 188743768 'pjTaskNumber2 -- 188743768 -- Represents the Number2 (Task) field.
    myArray(1168) = 188743996 'pjTaskNumber20 -- 188743996 -- Represents the Number20 (Task) field.
    myArray(1169) = 188743769 'pjTaskNumber3 -- 188743769 -- Represents the Number3 (Task) field.
    myArray(1170) = 188743770 'pjTaskNumber4 -- 188743770 -- Represents the Number4 (Task) field.
    myArray(1171) = 188743771 'pjTaskNumber5 -- 188743771 -- Represents the Number5 (Task) field.
    myArray(1172) = 188743982 'pjTaskNumber6 -- 188743982 -- Represents the Number6 (Task) field.
    myArray(1173) = 188743983 'pjTaskNumber7 -- 188743983 -- Represents the Number7 (Task) field.
    myArray(1174) = 188743984 'pjTaskNumber8 -- 188743984 -- Represents the Number8 (Task) field.
    myArray(1175) = 188743985 'pjTaskNumber9 -- 188743985 -- Represents the Number9 (Task) field.
    myArray(1176) = 188743777 'pjTaskObjects -- 188743777 -- Represents the Objects (Task) field.
    myArray(1177) = 188744096 'pjTaskOutlineCode1 -- 188744096 -- Represents the Outline Code1 (Task) field.
    myArray(1178) = 188744114 'pjTaskOutlineCode10 -- 188744114 -- Represents the Outline Code10 (Task) field.
    myArray(1179) = 188744098 'pjTaskOutlineCode2 -- 188744098 -- Represents the Outline Code2 (Task) field.
    myArray(1180) = 188744100 'pjTaskOutlineCode3 -- 188744100 -- Represents the Outline Code3 (Task) field.
    myArray(1181) = 188744102 'pjTaskOutlineCode4 -- 188744102 -- Represents the Outline Code4 (Task) field.
    myArray(1182) = 188744104 'pjTaskOutlineCode5 -- 188744104 -- Represents the Outline Code5 (Task) field.
    myArray(1183) = 188744106 'pjTaskOutlineCode6 -- 188744106 -- Represents the Outline Code6 (Task) field.
    myArray(1184) = 188744108 'pjTaskOutlineCode7 -- 188744108 -- Represents the Outline Code7 (Task) field.
    myArray(1185) = 188744110 'pjTaskOutlineCode8 -- 188744110 -- Represents the Outline Code8 (Task) field.
    myArray(1186) = 188744112 'pjTaskOutlineCode9 -- 188744112 -- Represents the Outline Code9 (Task) field.
    myArray(1187) = 188743765 'pjTaskOutlineLevel -- 188743765 -- Represents the Outline Level (Task) field.
    myArray(1188) = 188743782 'pjTaskOutlineNumber -- 188743782 -- Represents the Outline Code (Task) field.
    myArray(1189) = 188743905 'pjTaskOverallocated -- 188743905 -- Represents the Overallocated (Task) field.
    myArray(1190) = 188743848 'pjTaskOvertimeCost -- 188743848 -- Represents the Overtime Cost (Task) field.
    myArray(1191) = 188743843 'pjTaskOvertimeWork -- 188743843 -- Represents the Overtime Work (Task) field.
    myArray(1192) = 188745063 'pjTaskPathDrivenSuccessor -- 188745063 -- The driven successor field for a task path. See the�Task.PathDrivenSuccessor�property.
    myArray(1193) = 188745061 'pjTaskPathDrivingPredecessor -- 188745061 -- The driving predecessor field for a task path. See the�Task.PathDrivingPredecessor�property.
    myArray(1194) = 188745062 'pjTaskPathPredecessor -- 188745062 -- The predecessor field for a task path. See the�Task.PathPredecessor�property.
    myArray(1195) = 188745064 'pjTaskPathSuccessor -- 188745064 -- The successor field for a task path. See the�Task.PathSuccessor�property.
    myArray(1196) = 188743815 'pjTaskParentTask -- 188743815 -- Represents the Parent Task (Task) field.
    myArray(1197) = 188743712 'pjTaskPercentComplete -- 188743712 -- Represents the Percent Complete (Task) field.
    myArray(1198) = 188743713 'pjTaskPercentWorkComplete -- 188743713 -- Represents the Percent Work Complete (Task) field.
    myArray(1199) = 188744799 'pjTaskPhysicalPercentComplete -- 188744799 -- Represents the Physical Percent Complete (Task) field.
    myArray(1200) = 188744961 'pjTaskPlaceholder -- 188744961 -- Represents the Task Placeholder (Task) field.
    myArray(1201) = 188743727 'pjTaskPredecessors -- 188743727 -- Represents the Predecessors (Task) field.
    myArray(1202) = 188744050 'pjTaskPreleveledFinish -- 188744050 -- Represents the Preleveled Finish (Task) field.
    myArray(1203) = 188744049 'pjTaskPreleveledStart -- 188744049 -- Represents the Preleveled Start (Task) field.
    myArray(1204) = 188743705 'pjTaskPriority -- 188743705 -- Represents the Priority (Task) field.
    myArray(1205) = 188743764 'pjTaskProject -- 188743764 -- Represents the Project (Task) field.
    myArray(1206) = 188744930 'pjTaskRecalcFlags -- 188744930 -- Represents the Recalc Flags (Task) field.
    myArray(1207) = 188743809 'pjTaskRecurring -- 188743809 -- Represents the Recurring (Task) field.
    myArray(1208) = 188743846 'pjTaskRegularWork -- 188743846 -- Represents the Regular Work (Task) field.
    myArray(1209) = 188743690 'pjTaskRemainingCost -- 188743690 -- Represents the Remaining Cost (Task) field.
    myArray(1210) = 188743711 'pjTaskRemainingDuration -- 188743711 -- Represents the Remaining Duration (Task) field.
    myArray(1211) = 188743850 'pjTaskRemainingOvertimeCost -- 188743850 -- Represents the Remaining Overtime Cost (Task) field.
    myArray(1212) = 188743845 'pjTaskRemainingOvertimeWork -- 188743845 -- Represents the Remaining Overtime Work (Task) field.
    myArray(1213) = 188743684 'pjTaskRemainingWork -- 188743684 -- Represents the Remaining Work (Task) field.
    myArray(1214) = 188744809 'pjTaskResourceEnterpriseMultiValueCode20 -- 188744809 -- Represents the Enterprise Resource MultiValue Code20 (Task) field.
    myArray(1215) = 188744810 'pjTaskResourceEnterpriseMultiValueCode21 -- 188744810 -- Represents the Enterprise Resource MultiValue Code21 (Task) field.
    myArray(1216) = 188744811 'pjTaskResourceEnterpriseMultiValueCode22 -- 188744811 -- Represents the Enterprise Resource MultiValue Code22 (Task) field.
    myArray(1217) = 188744812 'pjTaskResourceEnterpriseMultiValueCode23 -- 188744812 -- Represents the Enterprise Resource MultiValue Code23 (Task) field.
    myArray(1218) = 188744813 'pjTaskResourceEnterpriseMultiValueCode24 -- 188744813 -- Represents the Enterprise Resource MultiValue Code24 (Task) field.
    myArray(1219) = 188744814 'pjTaskResourceEnterpriseMultiValueCode25 -- 188744814 -- Represents the Enterprise Resource MultiValue Code25 (Task) field.
    myArray(1220) = 188744815 'pjTaskResourceEnterpriseMultiValueCode26 -- 188744815 -- Represents the Enterprise Resource MultiValue Code26 (Task) field.
    myArray(1221) = 188744816 'pjTaskResourceEnterpriseMultiValueCode27 -- 188744816 -- Represents the Enterprise Resource MultiValue Code27 (Task) field.
    myArray(1222) = 188744817 'pjTaskResourceEnterpriseMultiValueCode28 -- 188744817 -- Represents the Enterprise Resource MultiValue Code28 (Task) field.
    myArray(1223) = 188744818 'pjTaskResourceEnterpriseMultiValueCode29 -- 188744818 -- Represents the Enterprise Resource MultiValue Code29 (Task) field.
    myArray(1224) = 188744769 'pjTaskResourceEnterpriseOutlineCode1 -- 188744769 -- Represents the Enterprise Resource Outline Code1 (Task) field.
    myArray(1225) = 188744778 'pjTaskResourceEnterpriseOutlineCode10 -- 188744778 -- Represents the Enterprise Resource Outline Code10 (Task) field.
    myArray(1226) = 188744779 'pjTaskResourceEnterpriseOutlineCode11 -- 188744779 -- Represents the Enterprise Resource Outline Code11 (Task) field.
    myArray(1227) = 188744780 'pjTaskResourceEnterpriseOutlineCode12 -- 188744780 -- Represents the Enterprise Resource Outline Code12 (Task) field.
    myArray(1228) = 188744781 'pjTaskResourceEnterpriseOutlineCode13 -- 188744781 -- Represents the Enterprise Resource Outline Code13 (Task) field.
    myArray(1229) = 188744782 'pjTaskResourceEnterpriseOutlineCode14 -- 188744782 -- Represents the Enterprise Resource Outline Code14 (Task) field.
    myArray(1230) = 188744783 'pjTaskResourceEnterpriseOutlineCode15 -- 188744783 -- Represents the Enterprise Resource Outline Code15 (Task) field.
    myArray(1231) = 188744784 'pjTaskResourceEnterpriseOutlineCode16 -- 188744784 -- Represents the Enterprise Resource Outline Code16 (Task) field.
    myArray(1232) = 188744785 'pjTaskResourceEnterpriseOutlineCode17 -- 188744785 -- Represents the Enterprise Resource Outline Code17 (Task) field.
    myArray(1233) = 188744786 'pjTaskResourceEnterpriseOutlineCode18 -- 188744786 -- Represents the Enterprise Resource Outline Code18 (Task) field.
    myArray(1234) = 188744787 'pjTaskResourceEnterpriseOutlineCode19 -- 188744787 -- Represents the Enterprise Resource Outline Code19 (Task) field.
    myArray(1235) = 188744770 'pjTaskResourceEnterpriseOutlineCode2 -- 188744770 -- Represents the Enterprise Resource Outline Code2 (Task) field.
    myArray(1236) = 188744788 'pjTaskResourceEnterpriseOutlineCode20 -- 188744788 -- Represents the Enterprise Resource Outline Code20 (Task) field.
    myArray(1237) = 188744789 'pjTaskResourceEnterpriseOutlineCode21 -- 188744789 -- Represents the Enterprise Resource Outline Code21 (Task) field.
    myArray(1238) = 188744790 'pjTaskResourceEnterpriseOutlineCode22 -- 188744790 -- Represents the Enterprise Resource Outline Code22 (Task) field.
    myArray(1239) = 188744791 'pjTaskResourceEnterpriseOutlineCode23 -- 188744791 -- Represents the Enterprise Resource Outline Code23 (Task) field.
    myArray(1240) = 188744792 'pjTaskResourceEnterpriseOutlineCode24 -- 188744792 -- Represents the Enterprise Resource Outline Code24 (Task) field.
    myArray(1241) = 188744793 'pjTaskResourceEnterpriseOutlineCode25 -- 188744793 -- Represents the Enterprise Resource Outline Code25 (Task) field.
    myArray(1242) = 188744794 'pjTaskResourceEnterpriseOutlineCode26 -- 188744794 -- Represents the Enterprise Resource Outline Code26 (Task) field.
    myArray(1243) = 188744795 'pjTaskResourceEnterpriseOutlineCode27 -- 188744795 -- Represents the Enterprise Resource Outline Code27 (Task) field.
    myArray(1244) = 188744796 'pjTaskResourceEnterpriseOutlineCode28 -- 188744796 -- Represents the Enterprise Resource Outline Code28 (Task) field.
    myArray(1245) = 188744797 'pjTaskResourceEnterpriseOutlineCode29 -- 188744797 -- Represents the Enterprise Resource Outline Code29 (Task) field.
    myArray(1246) = 188744771 'pjTaskResourceEnterpriseOutlineCode3 -- 188744771 -- Represents the Enterprise Resource Outline Code3 (Task) field.
    myArray(1247) = 188744772 'pjTaskResourceEnterpriseOutlineCode4 -- 188744772 -- Represents the Enterprise Resource Outline Code4 (Task) field.
    myArray(1248) = 188744773 'pjTaskResourceEnterpriseOutlineCode5 -- 188744773 -- Represents the Enterprise Resource Outline Code5 (Task) field.
    myArray(1249) = 188744774 'pjTaskResourceEnterpriseOutlineCode6 -- 188744774 -- Represents the Enterprise Resource Outline Code6 (Task) field.
    myArray(1250) = 188744775 'pjTaskResourceEnterpriseOutlineCode7 -- 188744775 -- Represents the Enterprise Resource Outline Code7 (Task) field.
    myArray(1251) = 188744776 'pjTaskResourceEnterpriseOutlineCode8 -- 188744776 -- Represents the Enterprise Resource Outline Code8 (Task) field.
    myArray(1252) = 188744777 'pjTaskResourceEnterpriseOutlineCode9 -- 188744777 -- Represents the Enterprise Resource Outline Code9 (Task) field.
    myArray(1253) = 188744798 'pjTaskResourceEnterpriseRBS -- 188744798 -- Represents the Enterprise Resource RBS (Task) field.
    myArray(1254) = 188743793 'pjTaskResourceGroup -- 188743793 -- Represents the Resource Group (Task) field.
    myArray(1255) = 188743730 'pjTaskResourceInitials -- 188743730 -- Represents the Resource Initials (Task) field.
    myArray(1256) = 188743729 'pjTaskResourceNames -- 188743729 -- Represents the Resource Names (Task) field.
    myArray(1257) = 188744029 'pjTaskResourcePhonetics -- 188744029 -- Represents the Resource Phonetics (Task) field.
    myArray(1258) = 188744131 'pjTaskResourceType -- 188744131 -- Represents the Resource Type (Task) field.
    myArray(1259) = 188743930 'pjTaskResponsePending -- 188743930 -- Represents the Response Pending (Task) field.
    myArray(1260) = 188743779 'pjTaskResume -- 188743779 -- Represents the Resume (Task) field.
    myArray(1261) = 188743781 'pjTaskResumeNoEarlierThan -- 188743781 -- Represents the Resume No Earlier Than (Task) field.
    myArray(1262) = 188743762 'pjTaskRollup -- 188743762 -- Represents the Rollup (Task) field.
    myArray(1263) = 188745020 'pjTaskScheduledDuration -- 188745020 -- Represents the Task Scheduled Duration (Task) field.
    myArray(1264) = 188745019 'pjTaskScheduledFinish -- 188745019 -- Represents the Task Scheduled Finish (Task) field.
    myArray(1265) = 188745018 'pjTaskScheduledStart -- 188745018 -- Represents the Task Scheduled Start (Task) field.
    myArray(1266) = 188743774 'pjTaskSheetNotes -- 188743774 -- Represents the Sheet Notes (Task) field.
    myArray(1267) = 188744218 'pjTaskSPI -- 188744218 -- Represents the SPI (Task) field.
    myArray(1268) = 188743715 'pjTaskStart -- 188743715 -- Represents the Start (Task) field.
    myArray(1269) = 188743732 'pjTaskStart1 -- 188743732 -- Represents the Start1 (Task) field.
    myArray(1270) = 188743970 'pjTaskStart10 -- 188743970 -- Represents the Start10 (Task) field.
    myArray(1271) = 188743735 'pjTaskStart2 -- 188743735 -- Represents the Start2 (Task) field.
    myArray(1272) = 188743738 'pjTaskStart3 -- 188743738 -- Represents the Start3 (Task) field.
    myArray(1273) = 188743741 'pjTaskStart4 -- 188743741 -- Represents the Start4 (Task) field.
    myArray(1274) = 188743744 'pjTaskStart5 -- 188743744 -- Represents the Start5 (Task) field.
    myArray(1275) = 188743962 'pjTaskStart6 -- 188743962 -- Represents the Start6 (Task) field.
    myArray(1276) = 188743964 'pjTaskStart7 -- 188743964 -- Represents the Start7 (Task) field.
    myArray(1277) = 188743966 'pjTaskStart8 -- 188743966 -- Represents the Start8 (Task) field.
    myArray(1278) = 188743968 'pjTaskStart9 -- 188743968 -- Represents the Start9 (Task) field.
    myArray(1279) = 188744118 'pjTaskStartSlack -- 188744118 -- Represents the Start Slack (Task) field.
    myArray(1280) = 188744965 'pjTaskStartText -- 188744965 -- Represents the Task Start Text (Task) field.
    myArray(1281) = 188743725 'pjTaskStartVariance -- 188743725 -- Represents the Start Variance (Task) field.
    myArray(1282) = 188744223 'pjTaskStatus -- 188744223 -- Represents the Status (Task) field.
    myArray(1283) = 188744801 'pjTaskStatusIndicator -- 188744801 -- Represents the Status Indicator (Task) field.
    myArray(1284) = 188744846 'pjTaskStatusManagerName -- 188744846 -- Represents the Status Manager Name (Task) field.
    myArray(1285) = 188743780 'pjTaskStop -- 188743780 -- Represents the Stop (Task) field.
    myArray(1286) = 188743706 'pjTaskSubproject -- 188743706 -- Represents the Subproject (Task) field.
    myArray(1287) = 188743926 'pjTaskSubprojectReadOnly -- 188743926 -- Represents the Subproject Read Only (Task) field.
    myArray(1288) = 188743728 'pjTaskSuccessors -- 188743728 -- Represents the Successors (Task) field.
    myArray(1289) = 188743772 'pjTaskSummary -- 188743772 -- Represents the Summary (Task) field.
    myArray(1290) = 188743693 'pjTaskSV -- 188743693 -- Represents the SV (Task) field.
    myArray(1291) = 188744220 'pjTaskSVPercent -- 188744220 -- Represents the SV Percent (Task) field.
    myArray(1292) = 188744222 'pjTaskTCPI -- 188744222 -- Represents the TCPI (Task) field.
    myArray(1293) = 188743931 'pjTaskTeamStatusPending -- 188743931 -- Represents the Team Status Pending (Task) field.
    myArray(1294) = 188743731 'pjTaskText1 -- 188743731 -- Represents the Text1 (Task) field.
    myArray(1295) = 188743750 'pjTaskText10 -- 188743750 -- Represents the Text10 (Task) field.
    myArray(1296) = 188743997 'pjTaskText11 -- 188743997 -- Represents the Text11 (Task) field.
    myArray(1297) = 188743998 'pjTaskText12 -- 188743998 -- Represents the Text12 (Task) field.
    myArray(1298) = 188743999 'pjTaskText13 -- 188743999 -- Represents the Text13 (Task) field.
    myArray(1299) = 188744000 'pjTaskText14 -- 188744000 -- Represents the Text14 (Task) field.
    myArray(1300) = 188744001 'pjTaskText15 -- 188744001 -- Represents the Text15 (Task) field.
    myArray(1301) = 188744002 'pjTaskText16 -- 188744002 -- Represents the Text16 (Task) field.
    myArray(1302) = 188744003 'pjTaskText17 -- 188744003 -- Represents the Text17 (Task) field.
    myArray(1303) = 188744004 'pjTaskText18 -- 188744004 -- Represents the Text18 (Task) field.
    myArray(1304) = 188744005 'pjTaskText19 -- 188744005 -- Represents the Text19 (Task) field.
    myArray(1305) = 188743734 'pjTaskText2 -- 188743734 -- Represents the Text2 (Task) field.
    myArray(1306) = 188744006 'pjTaskText20 -- 188744006 -- Represents the Text20 (Task) field.
    myArray(1307) = 188744007 'pjTaskText21 -- 188744007 -- Represents the Text21 (Task) field.
    myArray(1308) = 188744008 'pjTaskText22 -- 188744008 -- Represents the Text22 (Task) field.
    myArray(1309) = 188744009 'pjTaskText23 -- 188744009 -- Represents the Text23 (Task) field.
    myArray(1310) = 188744010 'pjTaskText24 -- 188744010 -- Represents the Text24 (Task) field.
    myArray(1311) = 188744011 'pjTaskText25 -- 188744011 -- Represents the Text25 (Task) field.
    myArray(1312) = 188744012 'pjTaskText26 -- 188744012 -- Represents the Text26 (Task) field.
    myArray(1313) = 188744013 'pjTaskText27 -- 188744013 -- Represents the Text27 (Task) field.
    myArray(1314) = 188744014 'pjTaskText28 -- 188744014 -- Represents the Text28 (Task) field.
    myArray(1315) = 188744015 'pjTaskText29 -- 188744015 -- Represents the Text29 (Task) field.
    myArray(1316) = 188743737 'pjTaskText3 -- 188743737 -- Represents the Text3 (Task) field.
    myArray(1317) = 188744016 'pjTaskText30 -- 188744016 -- Represents the Text30 (Task) field.
    myArray(1318) = 188743740 'pjTaskText4 -- 188743740 -- Represents the Text4 (Task) field.
    myArray(1319) = 188743743 'pjTaskText5 -- 188743743 -- Represents the Text5 (Task) field.
    myArray(1320) = 188743746 'pjTaskText6 -- 188743746 -- Represents the Text6 (Task) field.
    myArray(1321) = 188743747 'pjTaskText7 -- 188743747 -- Represents the Text7 (Task) field.
    myArray(1322) = 188743748 'pjTaskText8 -- 188743748 -- Represents the Text8 (Task) field.
    myArray(1323) = 188743749 'pjTaskText9 -- 188743749 -- Represents the Text9 (Task) field.
    myArray(1324) = 188743702 'pjTaskTotalSlack -- 188743702 -- Represents the Total Slack (Task) field.
    myArray(1325) = 188743808 'pjTaskType -- 188743808 -- Represents the Type (Task) field.
    myArray(1326) = 188743766 'pjTaskUniqueID -- 188743766 -- Represents the Unique ID (Task) field.
    myArray(1327) = 188743775 'pjTaskUniquePredecessors -- 188743775 -- Represents the Unique Predecessors (Task) field.
    myArray(1328) = 188743776 'pjTaskUniqueSuccessors -- 188743776 -- Represents the Unique Successors (Task) field.
    myArray(1329) = 188743791 'pjTaskUpdateNeeded -- 188743791 -- Represents the Update Needed (Task) field.
    myArray(1330) = 188744121 'pjTaskVAC -- 188744121 -- Represents the VAC (Task) field.
    myArray(1331) = 188744962 'pjTaskWarning -- 188744962 -- Represents the Task Warning (Task) field.
    myArray(1332) = 188743696 'pjTaskWBS -- 188743696 -- Represents the WBS (Task) field.
    myArray(1333) = 188744129 'pjTaskWBSPredecessors -- 188744129 -- Represents the WBS Predecessors (Task) field.
    myArray(1334) = 188744130 'pjTaskWBSSuccessors -- 188744130 -- Represents the WBS Successors (Task) field.
    myArray(1335) = 188743680 'pjTaskWork -- 188743680 -- Represents the Work (Task) field.
    myArray(1336) = 188743936 'pjTaskWorkContour -- 188743936 -- Represents the Work Contour (Task) field.
    myArray(1337) = 188743683 'pjTaskWorkVariance -- 188743683 -- Represents the Work Variance (Task) field.
    
    
    
    DebugEnum = myArray
End Function

