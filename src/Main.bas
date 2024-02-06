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
        Set Game = Tetris.GetTetris(ThisWorkbook.Worksheets("Tetris"), "K8")
        Call Game.Start
        Set Game = Nothing
    ElseIf Game.State > GameState.Paused Then
        Let Game.State = GameState.Paused
    ElseIf Game.State = GameState.Paused Then
        Let Game.State = GameState.Resumed
    End If
End Sub
Public Sub StopButton_Click()
With ThisWorkbook.Worksheets("Tetris")
    If Game Is Nothing Then
        Let .Range("B29").Value2 = "Not Started"
    Else
        Let Game.State = GameState.EndManually
        End
    End If
    
    Let .Range("B2").Value2 = "S T A R T"
End With
End Sub
