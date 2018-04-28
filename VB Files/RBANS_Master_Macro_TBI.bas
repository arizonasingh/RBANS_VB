Attribute VB_Name = "Master_Macro_TBI"
Sub RBANS_Master_Macro_TBI()

Worksheets("Raw_Data").Activate
Dim age As Integer
age = Range("B3")
Dim cell As Range

If age > 45 Then
 MsgBox "Invalid age for TBI Model study"
 Exit Sub
End If

If IsEmpty(Range("E12")) Then
    MsgBox "The program cannot run without valid Digit Span raw score"
    Exit Sub
ElseIf IsEmpty(Range("E13")) Then
    MsgBox "The program cannot run without valid Coding raw score"
    Exit Sub
End If

For Each cell In Range("E3:E4,E6:E7,E9:E10,E15:E18")
    If Not IsEmpty(cell) Then
        MsgBox "TBI Model RBANS only requires Digit Span and Coding"
        Exit Sub
    End If
Next

If age <= 19 Then
  Call RBANS_Form16_19
ElseIf age <=39 Then
  Call RBANS_Form20_39
ElseIf age <= 45 Then
  Call RBANS_Form40_49
End If

Worksheets("Raw_Data").Activate
Dim ID as Variant
ID = Application.InputBox("What is the Participant ID?", "Participant ID (number only)", 1)

Dim DSraw, DSscaled, COraw, COscaled, AttnIndex, AttnPerc As Integer
Dim AttnCI1, AttnCI2 As String

DSraw = Range("E12").Value
DSscaled = Range("G12").Value
COraw = Range("E13").Value
COscaled = Range("G13").Value
AttnIndex = Range("N2").Value
AttenPerc = Range("N4").Value

AttnCI1 = Left(Range("N3"), InStr(Range("N3"), "-") - 1)
AttnCI2 = Right(Range("N3"), InStr(Range("N3"), "-"))

Worksheets("Raw_Data").Activate
Range("E3:H4,E6:H7,E9:H10,E15:H18,F20,F22,K2:M4,O2:Q4").Select
Selection.ClearContents

Worksheets("TBI_Compiled_Data").Activate
Dim lastrow As long
lastrow = Cells.Find("*",SearchOrder:=xlByRows,SearchDirection:=xlPrevious).Row + 1

Range("A" & lastrow) = ID
Range("A" & lastrow+1) = ID & "--1"
Range("A" & lastrow+2) = ID & "--2"

Range("SF" & lastrow & ":SF" & lastrow+2) = 1
Range("SG" & lastrow & ":SG" & lastrow+2) = Application.InputBox("Examiner Initials","Examiner Initials")
Range("SH" & lastrow & ":SH" & lastrow+2) = DSraw
Range("SI" & lastrow & ":SI" & lastrow+2) = DSscaled
Range("SJ" & lastrow & ":SJ" & lastrow+2) = COraw
Range("SK" & lastrow & ":SK" & lastrow+2) = COscaled
Range("SL" & lastrow & ":SL" & lastrow+2) = AttnIndex
Range("SM" & lastrow & ":SM" & lastrow+2) = Abs(CInt(AttnCI1))
Range("SN" & lastrow & ":SN" & lastrow+2) = Abs(CInt(AttnCI2))
Range("SO" & lastrow & ":SO" & lastrow+2) = AttnPerc
Range("SP" & lastrow & ":SP" & lastrow+2) = 2

Worksheets("Raw_Data").Activate

End Sub
