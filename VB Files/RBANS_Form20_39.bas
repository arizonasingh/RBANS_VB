Attribute VB_Name = "form20_39"
Public Sub RBANS_Form20_39()

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

'Immediate Memory
  Worksheets("20_39").Activate
    Set Table = Range("C3:AB44")
    c = WorksheetFunction.Match(LL, Table.Columns(1))
    r = WorksheetFunction.Match(SM, Table.Rows(1))
    Intercept = Table.Cells(c, r).Value

  Worksheets("Raw_Data").Activate

  Range("F3:F4") = Intercept
  Range("K2") = Intercept
  IndexScale = Intercept

  If LL <= 13 Then
    SS1 = 1
  ElseIf LL <= 17 Then
    SS1 = 2
  ElseIf LL <= 19 Then
    SS1 = 3
  ElseIf LL <= 21 Then
    SS1 = 4
  ElseIf LL <= 23 Then
    SS1 = 5
  ElseIf LL = 24 Then
    SS1 = 6
  ElseIf LL <= 27 Then
    SS1 = 7
  ElseIf LL <= 29 Then
    SS1 = 8
  ElseIf LL <= 31 Then
    SS1 = 9
  ElseIf LL = 32 Then
    SS1 = 10
  ElseIf LL = 33 Then
    SS1 = 11
  ElseIf LL = 34 Then
    SS1 = 12
  ElseIf LL = 35 Then
    SS1 = 13
  ElseIf LL = 36 Then
    SS1 = 14
  ElseIf LL = 37 Then
    SS1 = 16
  ElseIf LL = 38 Then
    SS1 = 17
  ElseIf LL = 39 Then
    SS1 = 18
  ElseIf LL = 40 Then
    SS1 = 19
  End If

  If SM <= 6 Then
    SS2 = 1
  ElseIf SM <= 8 Then
    SS2 = 2
  ElseIf SM <= 10 Then
    SS2 = 3
  ElseIf SM <= 12 Then
    SS2 = 4
  ElseIf SM <= 14 Then
    SS2 = 5
  ElseIf SM = 15 Then
    SS2 = 6
  ElseIf SM <= 17 Then
    SS2 = 7
  ElseIf SM = 18 Then
    SS2 = 8
  ElseIf SM = 19 Then
    SS2 = 9
  ElseIf SM = 20 Then
    SS2 = 10
  ElseIf SM = 21 Then
    SS2 = 11
  ElseIf SM = 22 Then
    SS2 = 12
  ElseIf SM <= 23 Then
    SS2 = 14
  ElseIf SM <= 24 Then
    SS2 = 17
  End If

  Range("G3") = SS1
  Range("G4") = SS2

  CI1 = Intercept - 12
  CI2 = Intercept + 12
  Range("K3") = CI1 & "-" & CI2

  Worksheets("Index_Percentile_all").Activate
  rownum = Application.WorksheetFunction.Match(Intercept, Range("E:E"), 0)
  Percentile = Cells(rownum, 6)
  Worksheets("Raw_Data").Activate
  Range("K4") = Percentile

'Visuospatial/ Constructional
  Worksheets("20_39").Activate
    Set Table = Range("C48:X69")
    c = WorksheetFunction.Match(FC, Table.Columns(1))
    r = WorksheetFunction.Match(LO, Table.Rows(1))
    Intercept = Table.Cells(c, r).Value

  Worksheets("Raw_Data").Activate

  Range("F6:F7") = Intercept
  Range("L2") = Intercept
  IndexScale = IndexScale + Intercept

  If FC <= 15 Then
    SS1 = 1
  ElseIf FC = 16 Then
    SS1 = 2
  ElseIf FC = 17 Then
    SS1 = 4
  ElseIf FC = 18 Then
    SS1 = 6
  ElseIf FC = 19 Then
    SS1 = 9
  ElseIf FC = 20 Then
    SS1 = 12
  End If

  If LO <= 10 Then
    PG2 = "≤2"
  ElseIf LO <= 12 Then
    PG2 = "3-9"
  ElseIf LO = 13 Then
    PG2 = "10-16"
  ElseIf LO = 14 Then
    PG2 = "17-25"
  ElseIf LO <= 17 Then
    PG2 = "26-50"
  ElseIf LO <= 19 Then
    PG2 = "51-75"
  ElseIf LO = 20 Then
    PG2 = ">75"
  End If

  Range("G6") = SS1
  Range("H7") = PG2

  CI1 = Intercept - 14
  CI2 = Intercept + 14
  Range("L3") = CI1 & "-" & CI2

  Worksheets("Index_Percentile_all").Activate
  rownum = Application.WorksheetFunction.Match(Intercept, Range("E:E"), 0)
  Percentile = Cells(rownum, 6)
  Worksheets("Raw_Data").Activate
  Range("L4") = Percentile

'Language
  Worksheets("20_39").Activate
    Set Table = Range("C73:N114")
    c = WorksheetFunction.Match(SF, Table.Columns(1))
    r = WorksheetFunction.Match(PN, Table.Rows(1))
    Intercept = Table.Cells(c, r).Value

  Worksheets("Raw_Data").Activate

    Range("F9:F10") = Intercept
    Range("M2") = Intercept
    IndexScale = IndexScale + Intercept

    If PN <= 7 Then
      PG1 = "≤2"
    ElseIf PN = 8 Then
      PG1 = "3-9"
    ElseIf PN = 9 Then
      PG1 = "17-25"
    ElseIf PN = 10 Then
      PG1 = "51-75"
    End If

    If SF <= 12 Then
      SS2 = 1
    ElseIf SF <= 14 Then
      SS2 = 2
    ElseIf SF = 15 Then
      SS2 = 3
    ElseIf SF = 16 Then
      SS2 = 4
    ElseIf SF = 17 Then
      SS2 = 5
    ElseIf SF = 18 Then
      SS2 = 6
    ElseIf SF = 19 Then
      SS2 = 7
    ElseIf SF = 20 Then
      SS2 = 8
    ElseIf SF <= 22 Then
      SS2 = 10
    ElseIf SF <= 24 Then
      SS2 = 11
    ElseIf SF = 25 Then
      SS2 = 12
    ElseIf SF <= 27 Then
      SS2 = 13
    ElseIf SF <= 29 Then
      SS2 = 14
    ElseIf SF <= 31 Then
      SS2 = 15
    ElseIf SF <= 33 Then
      SS2 = 16
    ElseIf SF <= 35 Then
      SS2 = 17
    ElseIf SF <= 37 Then
      SS2 = 18
    ElseIf SF <= 40 Then
      SS2 = 19
    End If

    Range("H9") = PG1
    Range("G10") = SS2

    CI1 = Intercept - 15
    CI2 = Intercept + 15
    Range("M3") = CI1 & "-" & CI2

    Worksheets("Index_Percentile_all").Activate
    rownum = Application.WorksheetFunction.Match(Intercept, Range("E:E"), 0)
    Percentile = Cells(rownum, 6)
    Worksheets("Raw_Data").Activate
    Range("M4") = Percentile

'Attention
  Worksheets("20_39").Activate
    Set Table = Range("C118:T208")
    c = WorksheetFunction.Match(CO, Table.Columns(1))
    r = WorksheetFunction.Match(DS, Table.Rows(1))
    Intercept = Table.Cells(c, r).Value

  Worksheets("Raw_Data").Activate

  Range("F12:F13") = Intercept
  Range("N2") = Intercept
  IndexScale = IndexScale + Intercept

  If DS <= 5 Then
    SS1 = 1
  ElseIf DS = 6 Then
    SS1 = 2
  ElseIf DS = 7 Then
    SS1 = 4
  ElseIf DS = 8 Then
    SS1 = 5
  ElseIf DS = 9 Then
    SS1 = 6
  ElseIf DS = 10 Then
    SS1 = 8
  ElseIf DS = 11 Then
    SS1 = 9
  ElseIf DS = 12 Then
    SS1 = 10
  ElseIf DS = 13 Then
    SS1 = 11
  ElseIf DS = 14 Then
    SS1 = 13
  ElseIf DS = 15 Then
    SS1 = 14
  ElseIf DS = 16 Then
    SS1 = 15
  End If

  If CO <= 31 Then
    SS2 = 1
  ElseIf CO <= 34 Then
    SS2 = 2
  ElseIf CO <= 37 Then
    SS2 = 3
  ElseIf CO <= 40 Then
    SS2 = 4
  ElseIf CO <= 43 Then
    SS2 = 5
  ElseIf CO <= 46 Then
    SS2 = 6
  ElseIf CO <= 50 Then
    SS2 = 7
  ElseIf CO <= 53 Then
    SS2 = 8
  ElseIf CO <= 56 Then
    SS2 = 9
  ElseIf CO <= 59 Then
    SS2 = 10
  ElseIf CO <= 61 Then
    SS2 = 11
  ElseIf CO <= 64 Then
    SS2 = 12
  ElseIf CO <= 66 Then
    SS2 = 13
  ElseIf CO <= 68 Then
    SS2 = 14
  ElseIf CO <= 72 Then
    SS2 = 15
  ElseIf CO <= 74 Then
    SS2 = 16
  ElseIf CO <= 77 Then
    SS2 = 17
  ElseIf CO <= 81 Then
    SS2 = 18
  ElseIf CO <= 89 Then
    SS2 = 19
  End If

  Range("G12") = SS1
  Range("G13") = SS2

  CI1 = Intercept - 12
  CI2 = Intercept + 12
  Range("N3") = CI1 & "-" & CI2

  Worksheets("Index_Percentile_all").Activate
  rownum = Application.WorksheetFunction.Match(Intercept, Range("E:E"), 0)
  Percentile = Cells(rownum, 6)
  Worksheets("Raw_Data").Activate
  Range("N4") = Percentile

'Delayed Memory
  SumLSF = LLR + SMR + FCR
  Worksheets("20_39").Activate
    Set Table = Range("C212:X255")
    c = WorksheetFunction.Match(SumLSF, Table.Columns(1))
    r = WorksheetFunction.Match(LR, Table.Rows(1))
    Intercept = Table.Cells(c, r).Value

  Worksheets("Raw_Data").Activate

  Range("F15:F18") = Intercept
  Range("O2") = Intercept
  IndexScale = IndexScale + Intercept

  If LLR <= 3 Then
    PG1 = "≤2"
  ElseIf LLR = 4 Then
    PG1 = "3-9"
  ElseIf LLR = 5 Then
    PG1 = "10-16"
  ElseIf LLR <= 7 Then
    PG1 = "26-50"
  ElseIf LLR = 8 Then
    PG1 = "51-75"
  ElseIf LLR <= 10 Then
    PG1 = ">75"
  End If

  If LR <= 17 Then
    PG2 = "≤2"
  ElseIf LR = 18 Then
    PG2 = "3-9"
  ElseIf LR = 19 Then
    PG2 = "10-16"
  ElseIf LR = 20 Then
    PG2 = "51-75"
  End If

  If SMR <= 2 Then
    SS1 = 1
  ElseIf SMR = 3 Then
    SS1 = 2
  ElseIf SMR = 4 Then
    SS1 = 3
  ElseIf SMR = 5 Then
    SS1 = 4
  ElseIf SMR = 6 Then
    SS1 = 5
  ElseIf SMR = 7 Then
    SS1 = 6
  ElseIf SMR = 8 Then
    SS1 = 7
  ElseIf SMR = 9 Then
    SS1 = 8
  ElseIf SMR = 10 Then
    SS1 = 9
  ElseIf SMR = 11 Then
    SS1 = 11
  ElseIf SMR = 12 Then
    SS1 = 13
  End If

  If FCR <= 5 Then
    SS2 = 1
  ElseIf FCR <= 7 Then
    SS2 = 2
  ElseIf FCR <= 9 Then
    SS2 = 3
  ElseIf FCR = 10 Then
    SS2 = 4
  ElseIf FCR = 11 Then
    SS2 = 5
  ElseIf FCR = 12 Then
    SS2 = 6
  ElseIf FCR = 13 Then
    SS2 = 7
  ElseIf FCR = 14 Then
    SS2 = 8
  ElseIf FCR = 15 Then
    SS2 = 9
  ElseIf FCR = 16 Then
    SS2 = 10
  ElseIf FCR = 17 Then
    SS2 = 11
  ElseIf FCR = 18 Then
    SS2 = 12
  ElseIf FCR = 19 Then
    SS2 = 13
  ElseIf FCR = 20 Then
    SS2 = 15
  End If

  Range("H15") = PG1
  Range("H16") = PG2
  Range("G17") = SS1
  Range("G18") = SS2

  CI1 = Intercept - 12
  CI2 = Intercept + 12
  Range("O3") = CI1 & "-" & CI2

  Worksheets("Index_Percentile_all").Activate
  rownum = Application.WorksheetFunction.Match(Intercept, Range("E:E"), 0)
  Percentile = Cells(rownum, 6)
  Worksheets("Raw_Data").Activate
  Range("O4") = Percentile

'Total Scale
Range("F20") = IndexScale
Worksheets("Index_Percentile_all").Activate

rownum = Application.WorksheetFunction.Match(IndexScale, Range("A:A"), 0)
IndexScale = Cells(rownum, 2)
rownum = Application.WorksheetFunction.Match(IndexScale, Range("E:E"), 0)
Percentile = Cells(rownum, 6)

Worksheets("Raw_Data").Activate
Range("F22") = IndexScale
Range("Q2") = IndexScale
Range("Q4") = Percentile

CI1 = IndexScale - 8
CI2 = IndexScale + 8
Range("Q3") = CI1 & "-" & CI2

End Sub
