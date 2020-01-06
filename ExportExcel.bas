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
    
    Set excelapp = CreateObject("Excel.Application") 'for early binding: New Excel.Application
    excelapp.ScreenUpdating = False
    Set workbook = excelapp.Workbooks.Add()
    Set mySheet = workbook.Worksheets(1)
    
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
 
    mySheet.Cells(1, 1).Value = "�"
    mySheet.Cells(1, 2).Value = Application.FieldConstantToFieldName(PjField.pjTaskUniqueID) '"��� ���"
    mySheet.Cells(1, 3).Value = Application.FieldConstantToFieldName(PjField.pjTaskName) '"�������� ������"
    mySheet.Cells(1, 4).Value = Application.FieldConstantToFieldName(PjField.pjTaskStartText) '"������"
    mySheet.Cells(1, 5).Value = Application.FieldConstantToFieldName(PjField.pjTaskFinish)  '"���������"
    mySheet.Cells(1, 6).Value = Application.FieldConstantToFieldName(PjField.pjTaskDurationText) '"������������"
    mySheet.Cells(1, 7).Value = Application.FieldConstantToFieldName(PjField.pjTaskResourceNames) '"�������� ��������"
    mySheet.Cells(1, 8).Value = Application.FieldConstantToFieldName(PjField.pjTaskPredecessors)  '"���������������"
    

    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 1), mySheet.Cells(4, 1), True, 11, RGB(223, 227, 232))
    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 2), mySheet.Cells(4, 2), True, 11, RGB(223, 227, 232))
    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 3), mySheet.Cells(4, 3), True, 11, RGB(223, 227, 232))
    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 4), mySheet.Cells(4, 4), True, 11, RGB(223, 227, 232))
    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 5), mySheet.Cells(4, 5), True, 11, RGB(223, 227, 232))
    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 6), mySheet.Cells(4, 6), True, 11, RGB(223, 227, 232))
    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 7), mySheet.Cells(4, 7), True, 11, RGB(223, 227, 232))
    Call myFormat(excelapp, mySheet, mySheet.Cells(1, 8), mySheet.Cells(4, 8), True, 11, RGB(223, 227, 232))
    
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
            
            mySheet.Range(mySheet.Cells(2, 8 + myGanttMonthCount), mySheet.Cells(2, 8 + i - 1)).Select
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
            
            
            mySheet.Range(mySheet.Cells(1, 8 + myGanttYearCount), mySheet.Cells(1, 8 + i - 1)).Select
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
            
            mySheet.Range(mySheet.Cells(3, 8 + myGanttWeekCount), mySheet.Cells(3, 8 + i - 1)).Select
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
            mySheet.Range(mySheet.Cells(4, 8 + i), mySheet.Cells(4 + myActiveTaskCount, 8 + i)).Select
            excelapp.Selection.Interior.Color = RGB(223, 227, 232)
        End If
        
        mySheet.Cells(4, 8 + i).Value = Left(WeekdayName(Weekday(myDate, vbUseSystemDayOfWeek), True, vbUseSystemDayOfWeek), 1) 'was myDate
        mySheet.Cells(4, 8 + i).Select
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
            mySheet.Range(mySheet.Cells(2, 8 + myGanttMonthCount), mySheet.Cells(2, 8 + myLasti)).Select
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
            mySheet.Range(mySheet.Cells(1, 8 + myGanttYearCount), mySheet.Cells(1, 8 + myLasti)).Select
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
            mySheet.Range(mySheet.Cells(3, 8 + myGanttWeekCount), mySheet.Cells(3, 8 + myLasti)).Select
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
' Debug.Print "==================="
'  Debug.Print "name: """ & myTask.Name & """"
'  Debug.Print "unique ID: " & myTask.UniqueID
' Debug.Print "outline number: " & myTask.OutlineNumber
'  Debug.Print "outline level: " & myTask.OutlineLevel
'  Debug.Print "outline children count: " & myTask.OutlineChildren.Count
                
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
'   mySheet.Cells(currentLine, 8).Value = myTask.UniqueIDPredecessors
'    mySheet.Cells(currentLine, 9).Value = myTask.ID
                
                
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
                
'start Doing Gantt chart
'====================================================================
'tasks
                myColor = RGB(217, 225, 242)
                dateOffset = myD(myTask.Start) - myD(myStartDate)
                myDuration = myD(myTask.Finish) - myD(myTask.Start) + 1
                
                If myTask.Milestone Then
                    
                    mySheet.Cells(currentLine, 8 + dateOffset + 1).Value = ChrW(&H25CA)
                    mySheet.Cells(currentLine, 8 + dateOffset + 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
                    mySheet.Cells(currentLine, 8 + dateOffset + 1).VerticalAlignment = XlVAlign.xlVAlignCenter
                    
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
                   
                    Call GanttFormat(excelapp, mySheet, mySheet.Cells(currentLine, 8 + dateOffset + 1), mySheet.Cells(currentLine, 8 + dateOffset + myDuration), myColor)
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
                       
     For i = 1 To 8
         mySheet.Columns(i).AutoFit
         If mySheet.Columns(i).ColumnWidth > 100 Then
             mySheet.Columns(i).ColumnWidth = 100
         End If
    
     Next i
                            
    For i = 9 To numberOfDays + 9
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
            
' Debug.Print "==================="
'  Debug.Print "name: """ & myTask.Name & """"
' Debug.Print "unique ID: " & myTask.UniqueID
' Debug.Print "outline number: " & myTask.OutlineNumber
' Debug.Print "outline level: " & myTask.OutlineLevel
'  Debug.Print "outline children count: " & myTask.OutlineChildren.Count
            
            
            
'  myTask.ResourceNames
'  myTask.Predecessors
'  myTask.UniqueIDPredecessors
            
            
' myTask.OutlineNumber
' myTask.Name
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

