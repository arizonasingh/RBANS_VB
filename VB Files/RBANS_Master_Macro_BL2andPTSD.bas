Attribute VB_Name = "Master_Macro_BL2andPTSD"
Sub RBANS_Master_Macro_BL2andPTSD()

Worksheets("Raw_Data").Activate
Dim age As Integer
age = Range("B3")
Dim cell As Range

For Each cell In Range("E3:E4,E6:E7,E9:E10,E12:E13,E15:E18")
    If IsEmpty(cell) Then
        MsgBox "The program requires all subtests to have raw score entered!"
        Exit Sub
    End If
Next

If age <= 19 Then
  Call RBANS_Form16_19
ElseIf age <=39 Then
  Call RBANS_Form20_39
ElseIf age <= 49 Then
  Call RBANS_Form40_49
ElseIf age <= 50 Then
  Call RBANS_Form50_59
End If

End Sub
