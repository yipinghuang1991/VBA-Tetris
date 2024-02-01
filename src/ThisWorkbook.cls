VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    With ThisWorkbook.Worksheets("Tetris")
        Let .Range("U24").Value2 = vbNullString
        Let .Range("U26").Value2 = "S T A R T"
    End With
End Sub
