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
    ElseIf Game.GameState > 0 Then
        Let Game.GameState = 0
    ElseIf Game.GameState = 0 Then
        Let Game.GameState = 2
    End If
End Sub
'-----------------------------------------------------------------------------80
Public Sub StopButton_Click()
    If Game Is Nothing Then
        With ThisWorkbook.Worksheets("Tetris")
            Let .Range("U24").Value2 = "Not Started"
            Let .Range("U26").Value2 = "S T A R T"
        End With
    Else
        Let Game.GameState = -2
        End
    End If
End Sub

