VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tetromino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'=============================================================================80
' The Tetromino class holds the shape, color and stance of the piece,
' its center position, and the blocks that compose the piece.
'
' Author: Yi-Ping Huang
' https://github.com/yipinghuang1991/VBA-Tetris
'-----------------------------------------------------------------------------80
' In my implementation, top left most cell of the board has coordinate (1, 1).
' +X direction is pointing downwards, +Y direction is pointing right.
' That is, the board has coordinate just like a matrix in the mathematical sense.
' By the right hand rule, +Z direction is pointing out of the page (screen?).
' In this convention, to rotate COUNTERCLOCKWISE is to rotating positively.
' Let the default spawn state be of stance 0.
' Rotate CCW once be of stance 1, since it rotates 1 * (pi/2) radians.
' Rotate CCW once more, the stance would be 2, because it rotates 2 * (pi/2) radians.
' Likewise, rotate CCW once more, the stance would be 3.
' Finally, rotate CCW the fourth time, the stance would be 0 (= 4 mod 4).
'=============================================================================80
' Enumeration
'=============================================================================80
Public Enum TetrominoShape
    IShape = 1
    OShape
    TShape
    JShape
    LShape
    SShape
    ZShape
End Enum
'=============================================================================80
' Variable Container
'=============================================================================80
Private Type Variable
    x As Long
    y As Long
    Stance As Long
    Color As Long
    Blocks() As Block
End Type
'-----------------------------------------------------------------------------80
Private This As Variable
'=============================================================================80
' Event
'=============================================================================80
Private Sub Class_Initialize()
    ReDim Blocks(1 To 4)
End Sub
'=============================================================================80
' Read-Only Property
'=============================================================================80
Public Property Get Self() As Tetromino
    Set Self = Me
End Property
'=============================================================================80
' Read-Write Property
'=============================================================================80
Public Property Get x() As Long
    Let x = This.x
End Property
Public Property Let x(ByVal x As Long)
    Let This.x = x
End Property
'-----------------------------------------------------------------------------80
Public Property Get y() As Long
    Let y = This.y
End Property
Public Property Let y(ByVal y As Long)
    Let This.y = y
End Property
'-----------------------------------------------------------------------------80
Public Property Get Stance() As Long
    Let Stance = This.Stance
End Property
Public Property Let Stance(ByVal Stance As Long)
    Let This.Stance = (4 + Stance) Mod 4 ' Correction to VBA's modulo operation
End Property
'-----------------------------------------------------------------------------80
Public Property Get Color() As Long
    Let Color = This.Color
End Property
Public Property Let Color(ByVal Color As Long)
    Let This.Color = Color
End Property
'-----------------------------------------------------------------------------80
Public Property Get Blocks() As Block()
    Let Blocks = This.Blocks
End Property
Public Property Let Blocks(ByRef Blocks() As Block)
    Let This.Blocks = Blocks
End Property
'=============================================================================80
' Constructor
'=============================================================================80
Public Function GetTetromino(ByVal Shape As TetrominoShape, ByVal x As Long, ByVal y As Long) As Tetromino
With New Tetromino
    Let .x = x
    Let .y = y
    Let .Color = GetTetrominoColor(Shape)
    Let .Blocks = GetTetrominoBlocks(Shape)
    
    Dim B As Variant
    For Each B In .Blocks
        Set B.Parent = .Self
    Next B
    
    If Shape = LShape Or Shape = ZShape Then
        Call .Reflect
    End If
    
    Set GetTetromino = .Self
End With
End Function
'-----------------------------------------------------------------------------80
Public Function Copy() As Tetromino
With New Tetromino
    Let .x = This.x
    Let .y = This.y
    Let .Stance = This.Stance
    Let .Color = This.Color
    
    Dim Blocks() As Block: ReDim Blocks(1 To 4)
    
    Dim i As Long
    For i = 1 To 4
        Set Blocks(i) = This.Blocks(i).Copy
        Set Blocks(i).Parent = .Self
    Next i
    
    Let .Blocks = Blocks
    
    Set Copy = .Self
End With
End Function
'=============================================================================80
' Method
'=============================================================================80
Public Sub Translate(ByVal x As Long, ByVal y As Long)
With This
    Let .x = x + .x
    Let .y = y + .y
End With
End Sub
'-----------------------------------------------------------------------------80
Public Sub Rotate(Optional ByVal Sign As Long = -1) ' CCW: Sign = 1, CW: Sign = -1
With This
    Dim B As Variant
    For Each B In .Blocks
        Call B.Rotate(Sign)
    Next B
    
    Let Stance = Sign + .Stance
End With
End Sub
'-----------------------------------------------------------------------------80
Public Sub Reflect(Optional ByVal x As Boolean = True, Optional ByVal y As Boolean)
    Dim B As Variant
    For Each B In This.Blocks
        Call B.Reflect(x, y)
    Next B
End Sub
'=============================================================================80
' Helper Function
'=============================================================================80
Private Function GetTetrominoBlocks(ByVal Shape As TetrominoShape) As Block()
    Dim Blocks() As Block: ReDim Blocks(1 To 4)
    ' The coordinates are with respect to their rotation center
    Select Case Shape
        Case IShape: Set Blocks(1) = Block.GetBlock(0, -1, True)
                     Set Blocks(2) = Block.GetBlock(0, 0, True)
                     Set Blocks(3) = Block.GetBlock(0, 1, True)
                     Set Blocks(4) = Block.GetBlock(0, 2, True)
        Case OShape: Set Blocks(1) = Block.GetBlock(0, 0, True)
                     Set Blocks(2) = Block.GetBlock(0, 1, True)
                     Set Blocks(3) = Block.GetBlock(1, 1, True)
                     Set Blocks(4) = Block.GetBlock(1, 0, True)
        Case TShape: Set Blocks(1) = Block.GetBlock(-1, 0)
                     Set Blocks(2) = Block.GetBlock(0, -1)
                     Set Blocks(3) = Block.GetBlock(0, 0)
                     Set Blocks(4) = Block.GetBlock(0, 1)
        Case JShape, LShape
                     Set Blocks(1) = Block.GetBlock(0, 1)
                     Set Blocks(2) = Block.GetBlock(0, 0)
                     Set Blocks(3) = Block.GetBlock(0, -1)
                     Set Blocks(4) = Block.GetBlock(-1, -1)
'        Case LShape: Let Blocks(1) = Block.GetBlock(0, -1)
'                     Let Blocks(2) = Block.GetBlock(0, 0)
'                     Let Blocks(3) = Block.GetBlock(0, 1)
'                     Let Blocks(4) = Block.GetBlock(-1, 1)
        Case SShape, ZShape
                     Set Blocks(1) = Block.GetBlock(-1, 1)
                     Set Blocks(2) = Block.GetBlock(-1, 0)
                     Set Blocks(3) = Block.GetBlock(0, 0)
                     Set Blocks(4) = Block.GetBlock(0, -1)
'        Case ZShape: Let Blocks(1) = Block.GetBlock(-1, -1)
'                     Let Blocks(2) = Block.GetBlock(-1, 0)
'                     Let Blocks(3) = Block.GetBlock(0, 0)
'                     Let Blocks(4) = Block.GetBlock(0, 1)
    End Select
    
    Let GetTetrominoBlocks = Blocks
End Function
'-----------------------------------------------------------------------------80
Private Function GetTetrominoColor(ByVal Shape As TetrominoShape) As Long
Select Case Shape
    Case IShape: Let GetTetrominoColor = RGB(0, 255, 255)
    Case OShape: Let GetTetrominoColor = RGB(255, 255, 0)
    Case TShape: Let GetTetrominoColor = RGB(128, 0, 128)
    Case JShape: Let GetTetrominoColor = RGB(0, 0, 255)
    Case LShape: Let GetTetrominoColor = RGB(255, 128, 0)
    Case SShape: Let GetTetrominoColor = RGB(0, 255, 0)
    Case ZShape: Let GetTetrominoColor = RGB(255, 0, 0)
End Select
End Function
