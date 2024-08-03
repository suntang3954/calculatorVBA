VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 计算器 
   Caption         =   "UserForm1"
   ClientHeight    =   6888
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4800
   OleObjectBlob   =   "计算器.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "计算器"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim expression As String

Private Sub btn0_Click()
    'add the new number
    addNewNumber "0"
End Sub
Private Sub btn1_Click()
    addNewNumber "1"
End Sub
Private Sub btn2_Click()
    addNewNumber "2"
End Sub
Private Sub btn3_Click()
    addNewNumber "3"
End Sub
Private Sub btn4_Click()
    addNewNumber "4"
End Sub
Private Sub btn5_Click()
    addNewNumber "5"
End Sub
Private Sub btn6_Click()
    addNewNumber "6"
End Sub
Private Sub btn7_Click()
    addNewNumber "7"
End Sub
Private Sub btn8_Click()
    addNewNumber "8"
End Sub
Private Sub btn9_Click()
    addNewNumber "9"
End Sub
Private Sub btnAC_Click()
    '获取计算结果控件的值
    Dim result As String
    result = Me.ResultBar.Text
    '删除最后一位字符
    If Len(result) > 0 Then
        Me.ResultBar.Value = Left(result, Len(result) - 1)
    Else
        Me.ResultBar.Value = ""
    End If
End Sub

Private Sub btnAC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.ResultBar.Value = ""
End Sub

Private Sub btnEqual_Click()
    Dim expression As String
    Dim result As Double

    expression = Me.ResultBar.Text
    On Error Resume Next
    result = Evaluate(expression)
    If Err.Number <> 0 Then
        MsgBox "Invalid Expression", vbExclamation, "Invalid"
        Me.ResultBar.Text = ""
    Else
        Me.ResultBar.Text = result
    End If
    On Error GoTo 0
End Sub

Private Sub btnPlus_Click()
    Me.ResultBar = Me.ResultBar & " + "
End Sub
Private Sub btnDivde_Click()
    Me.ResultBar = Me.ResultBar & " / "
End Sub

Private Sub btnMinus_Click()
    Me.ResultBar = Me.ResultBar & " - "
End Sub

Private Sub btnMultipe_Click()
    Me.ResultBar = Me.ResultBar & " * "
End Sub

Private Sub btnPercent_Click()
    Me.ResultBar = Me.ResultBar.Value / 100 & "%"
End Sub
Private Sub btnShift_Click()
    Me.ResultBar = -1 * Me.ResultBar.Value
End Sub

Private Sub ResultBar_Change()

End Sub

Private Sub UserForm_Click()
    
End Sub

Function addNewNumber(newNumber As String)
    If Me.ResultBar.Text Like "*[+/-]" Then
        Me.ResultBar.Text = Mid(Me.ResultBar.Text, 1, InStrRev(Me.ResultBar.Text, " "))
    End If
    Me.ResultBar.Text = Me.ResultBar.Text & newNumber
End Function
