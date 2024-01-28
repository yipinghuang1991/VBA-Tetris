Attribute VB_Name = "Main"
Option Explicit
'=============================================================================80
' For buttons
'
' Author: Yi-Ping Huang
' https://github.com/yipinghuang1991/VBA-Tetris
'=============================================================================80
'=============================================================================80
' Variable
'=============================================================================80
Private Game As Tetris
'=============================================================================80
' Method
'=============================================================================80
Public Sub StartButton_Click()
    If Game Is Nothing Then
        Set Game = Tetris.GetTetris(ThisWorkbook.Worksheets("Tetris"), "I8")
        Call Game.Start
        Set Game = Nothing
    Else
        Let Game.GameState = -2
        ' Set Game = Nothing ' Why don't I need this line?
    End If
End Sub
