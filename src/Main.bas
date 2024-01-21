Attribute VB_Name = "Main"
Option Explicit
'=============================================================================80
' Tetris
' Author: Yi-Ping Huang
' https://github.com/yipinghuang1991/VBA-Tetris
'=============================================================================80
'=============================================================================80
' Variable
'=============================================================================80
Private Tet As Tetris
Private HasGameStarted As Boolean
Private IsGamePaused As Boolean
'=============================================================================80
' Methods
'=============================================================================80
Public Sub Start()
    If Not HasGameStarted Then
        Set Tet = Tetris.GetTetris(ThisWorkbook.Worksheets("Tetris"), "J8")
        Let HasGameStarted = True
        Call Tet.Start
    End If
End Sub
'-----------------------------------------------------------------------------80
Public Sub Terminate()
    Set Tet = Nothing
    End
End Sub
'=============================================================================80
' Helper Function
'=============================================================================80
Public Sub GravityDropMain()
    Call Tet.ResetGravity
End Sub
