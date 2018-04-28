Attribute VB_Name = "rbans_update_scoring"
Public Sub RBANS_update_scoring()

  Dim LL, SM, FC, LO, PN, SF, DS, CO, LLR, LR, SMR, FCR As Variant
  Dim Table As Range

  LL = Range("E3")  'list learning
  SM = Range("E4")  'story memory
  FC = Range("E6")  'figure copy
  LO = Range("E7")  'line orientation
  PN = Range("E9")  'picture naming
  SF = Range("E10") 'semantic fluency
  DS = Range("E12") 'digit span
  CO = Range("E13")  'coding
  LLR = Range("E15")  'list recall
  LR = Range("E16") 'list recognition
  SMR = Range("E17")  'story recall
  FCR = Range("E18")  'figure recall

  age = Range("B3")

  Dim r, c, SS1, SS2, PG1, PG2, SumLSF As Variant 'row, column, scaled score, percentile group, and sum of list recall, story recall, and figure recall
  Dim Intercept, CI1, CI2, IndexScale, Percentile, rownum As Variant

   If age <= 19 Then
      Call RBANS_Form16_19
   ElseIf age <= 39 Then
      Call RBANS_Form20_39
   ElseIf age <= 49 Then
      Call RBANS_Form40_49
   ElseIf age = 50 Then
      Call RBANS_Form50_59
   End If

End Sub
