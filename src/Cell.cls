VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=============================================================================80
' This is a wrapper for the actual cell range object of Excel worksheet.
' It's purpose is to determine if the cell should have change color.
'
' Author: Yi-Ping Huang
' https://github.com/yipinghuang1991/VBA-Tetris
'=============================================================================80
'=============================================================================80
' Variable Container
'=============================================================================80
Private Type Variable
    State As Long ' -1 = unoccupied, 0 = grounded block, 1 = part of piece
    BackgroundColor As Long
    Range As Range
    Block As Block
End Type
'-----------------------------------------------------------------------------80
Private This As Variable
'=============================================================================80
' Event
'=============================================================================80
Private Sub Class_Initialize()
With This
    Let .State = -1
End With
End Sub
'=============================================================================80
' Read-Only Property
'=============================================================================80
Public Property Get Self() As Cell
    Set Self = Me
End Property
'-----------------------------------------------------------------------------80
Public Property Get State() As Long
    Let State = This.State
End Property
'-----------------------------------------------------------------------------80
Public Property Get Color() As Long
With This
    Select Case .State
        Case -1:   Let Color = .BackgroundColor
        Case Else: Let Color = .Block.Color
    End Select
End With
End Property
'=============================================================================80
' Write-Only Property
'=============================================================================80
Public Property Let BackgroundColor(ByVal BackgroundColor As Long)
    Let This.BackgroundColor = BackgroundColor
End Property
'-----------------------------------------------------------------------------80
Public Property Set Range(ByRef Range As Range)
    Set This.Range = Range
End Property
'=============================================================================80
' Read-Write Property
'=============================================================================80
Public Property Get Block() As Block
    Set Block = This.Block
End Property
Public Property Set Block(ByRef Block As Block)
With This
    Select Case True
        Case Block Is Nothing: Let .State = -1
        Case Block.HasParent:  Let .State = 1
        Case Else:             Let .State = 0
    End Select
    
    Set .Block = Block
End With
End Property
'=============================================================================80
' Method
'=============================================================================80
Public Sub Draw()
    Let This.Range.Interior.Color = Color
End Sub
