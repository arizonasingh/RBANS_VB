Attribute VB_Name = "reset_table"
Public Sub reset_table()

Worksheets("Raw_Data").Activate
Range("B1,B3,E3:H4,E6:H7,E9:H10,E12:H13,E15:H18,F20,F22,K2:O4,Q2:Q4").Select
Range("B1").Activate
Selection.ClearContents

End Sub
