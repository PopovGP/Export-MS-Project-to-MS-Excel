'Attribute VB_Name = "ExportParametersForm"
'@Lang VBA
Option Explicit

Private Sub UpdateColors()
    Dim ctrl As Control
    Dim i As Integer
    
    For Each ctrl In Me.Controls
        ' Check if the control is a TextBox
        If TypeName(ctrl) = "TextBox" Then
           
            If ctrl.Text = "" Then
                ctrl.Text = "0"
            End If
            
            i = CInt(ctrl.Text)
                
            If i < 0 Then
                ctrl.Text = "0"
            End If
            
            If i > 255 Then
                ctrl.Text = "255"
            End If
                
        End If
    Next ctrl
    
    color0 = RGB(CInt(TextBoxR0.Text), CInt(TextBoxG0.Text), CInt(TextBoxB0.Text))
    LabelC0.BackColor = color0
    
    color1 = RGB(CInt(TextBoxR1.Text), CInt(TextBoxG1.Text), CInt(TextBoxB1.Text))
    LabelC1.BackColor = color1
 
    color2 = RGB(CInt(TextBoxR2.Text), CInt(TextBoxG2.Text), CInt(TextBoxB2.Text))
    LabelC2.BackColor = color2
 
    color3 = RGB(CInt(TextBoxR3.Text), CInt(TextBoxG3.Text), CInt(TextBoxB3.Text))
    LabelC3.BackColor = color3

    color4 = RGB(CInt(TextBoxR4.Text), CInt(TextBoxG4.Text), CInt(TextBoxB4.Text))
    LabelC4.BackColor = color4

    color5 = RGB(CInt(TextBoxR5.Text), CInt(TextBoxG5.Text), CInt(TextBoxB5.Text))
    LabelC5.BackColor = color5

End Sub

Private Sub CommandButton1_Click()
    Call UpdateColors
End Sub



Private Sub Frame2_Click()
 Call UpdateColors
End Sub

Private Sub OkButton_Click()
    Call UpdateColors
    myCommand = 1
    
    If OptButtonYMWD.Value = True Then
        myCommand = 1
    ElseIf OptButtonMonths = True Then
        myCommand = 2
    ElseIf OptButtonQuarterly = True Then
        myCommand = 3
    Else
        myCommand = 4
    End If
        
    Unload Me
    
    
End Sub





Private Sub TextBoxB0_Change()
    Call UpdateColors
End Sub

Private Sub TextBoxB1_Change()
    Call UpdateColors
End Sub

Private Sub TextBoxB2_Change()
    Call UpdateColors
End Sub

Private Sub TextBoxB4_Change()
    Call UpdateColors
End Sub

Private Sub TextBoxB5_Change()
    Call UpdateColors
End Sub

Private Sub TextBoxG0_Change()
    Call UpdateColors
End Sub

Private Sub TextBoxG1_Change()
    Call UpdateColors
End Sub

Private Sub TextBoxG2_Change()
    Call UpdateColors
End Sub

Private Sub TextBoxG3_Change()
    Call UpdateColors
End Sub

Private Sub TextBoxG4_Change()
    Call UpdateColors
End Sub

Private Sub TextBoxG5_Change()
    Call UpdateColors
End Sub

Private Sub TextBoxR0_Change()
    Call UpdateColors
End Sub

Private Sub TextBoxR1_Change()
    Call UpdateColors
End Sub

Private Sub TextBoxR2_Change()
    Call UpdateColors
End Sub

Private Sub TextBoxR3_Change()
    Call UpdateColors
End Sub

Private Sub TextBoxR4_Change()
    Call UpdateColors
End Sub

Private Sub TextBoxR5_Change()
    Call UpdateColors
End Sub

Private Sub UserForm_Initialize()
    myCommand = 0
    OptButtonYMWD.Value = True
    
    Call UpdateColors
End Sub

Private Sub CancelButton_Click()
    myCommand = 0
    Unload Me
End Sub


