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
'=============================================================================80
' Methods
'=============================================================================80
Public Sub ButtonClick()
    If Tet Is Nothing Then
        Set Tet = Tetris.GetTetris(ThisWorkbook.Worksheets("Tetris"), "J8")
        Call Tet.Start
    End If
End Sub
'-----------------------------------------------------------------------------80
Public Sub Terminate()
    Set Tet = Nothing
    End
End Sub
