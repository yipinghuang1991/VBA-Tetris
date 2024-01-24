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
' Tetris
' Author: Yi-Ping Huang
' https://github.com/yipinghuang1991/VBA-Tetris
'=============================================================================80
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
    X As Long
    Y As Long
    Stance As Long
    Color As Long
    Blocks As Collection
End Type
'-----------------------------------------------------------------------------80
Private This As Variable
'=============================================================================80
' Event
'=============================================================================80
Private Sub Class_Initialize()
    Set This.Blocks = New Collection
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
Public Property Get X() As Long
    Let X = This.X
End Property
Public Property Let X(ByVal X As Long)
    Let This.X = X
End Property
'-----------------------------------------------------------------------------80
Public Property Get Y() As Long
    Let Y = This.Y
End Property
Public Property Let Y(ByVal Y As Long)
    Let This.Y = Y
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
Public Property Get Blocks() As Collection
    Set Blocks = This.Blocks
End Property
Public Property Set Blocks(ByRef Blocks As Collection)
    Set This.Blocks = Blocks
End Property
'=============================================================================80
' Constructor
'=============================================================================80
Public Function GetTetromino(ByVal Shape As TetrominoShape, ByVal X As Long, ByVal Y As Long) As Tetromino
With New Tetromino
    Let .X = X
    Let .Y = Y
    Let .Color = GetTetrominoColor(Shape)
    Set .Blocks = GetTetrominoBlocks(Shape)
    Call .SetBlocksParent(.Self)
    
    If Shape = TetrominoShape.LShape Or Shape = TetrominoShape.ZShape Then
        Call .Reflect
    End If
    
    Set GetTetromino = .Self
End With
End Function
'-----------------------------------------------------------------------------80
Public Function Copy() As Tetromino
With New Tetromino
    Let .X = This.X
    Let .Y = This.Y
    Let .Stance = This.Stance
    Let .Color = This.Color
    
    Dim B As Block
    For Each B In This.Blocks
        Call .Blocks.Add(B.Copy)
    Next B
    
    For Each B In .Blocks
        Set B.Parent = .Self
    Next B
    
    Set Copy = .Self
End With
End Function
'=============================================================================80
' Method
'=============================================================================80
Public Sub SetBlocksParent(ByRef Parent As Tetromino)
    Dim B As Block
    For Each B In This.Blocks
        Set B.Parent = Parent
    Next B
End Sub
'-----------------------------------------------------------------------------80
Public Sub Translate(ByVal X As Long, ByVal Y As Long)
With This
    Let .X = X + .X
    Let .Y = Y + .Y
End With
End Sub
'-----------------------------------------------------------------------------80
Public Sub Rotate(Optional ByVal Sign As Long = -1) ' CCW: Sign = 1, CW: Sign = -1
With This
    Dim B As Block
    For Each B In .Blocks
        Call B.Rotate(Sign)
    Next B
    
    Let Stance = Sign + .Stance
End With
End Sub
'-----------------------------------------------------------------------------80
Public Sub Reflect(Optional ByVal X As Boolean = True, Optional ByVal Y As Boolean)
    Dim B As Block
    For Each B In This.Blocks
        Call B.Reflect(X, Y)
    Next B
End Sub
'=============================================================================80
' Helper Function
'=============================================================================80
Private Function GetTetrominoBlocks(ByVal Shape As TetrominoShape) As Collection
    Dim Coll As Collection: Set Coll = New Collection
    ' The coordinates are with respect to their rotation center
    Select Case Shape
        Case TetrominoShape.IShape: Call Coll.Add(Block.GetBlock(0, -1, True))
                                    Call Coll.Add(Block.GetBlock(0, 0, True))
                                    Call Coll.Add(Block.GetBlock(0, 1, True))
                                    Call Coll.Add(Block.GetBlock(0, 2, True))
        Case TetrominoShape.OShape: Call Coll.Add(Block.GetBlock(0, 0, True))
                                    Call Coll.Add(Block.GetBlock(0, 1, True))
                                    Call Coll.Add(Block.GetBlock(1, 1, True))
                                    Call Coll.Add(Block.GetBlock(1, 0, True))
        Case TetrominoShape.TShape: Call Coll.Add(Block.GetBlock(-1, 0))
                                    Call Coll.Add(Block.GetBlock(0, -1))
                                    Call Coll.Add(Block.GetBlock(0, 0))
                                    Call Coll.Add(Block.GetBlock(0, 1))
        Case TetrominoShape.JShape, TetrominoShape.LShape
                                    Call Coll.Add(Block.GetBlock(0, 1))
                                    Call Coll.Add(Block.GetBlock(0, 0))
                                    Call Coll.Add(Block.GetBlock(0, -1))
                                    Call Coll.Add(Block.GetBlock(-1, -1))
'        Case TetrominoShape.LShape: Call Coll.Add(Block.GetBlock(0, -1))
'                                    Call Coll.Add(Block.GetBlock(0, 0))
'                                    Call Coll.Add(Block.GetBlock(0, 1))
'                                    Call Coll.Add(Block.GetBlock(-1, 1))
        Case TetrominoShape.SShape, TetrominoShape.ZShape
                                    Call Coll.Add(Block.GetBlock(-1, 1))
                                    Call Coll.Add(Block.GetBlock(-1, 0))
                                    Call Coll.Add(Block.GetBlock(0, 0))
                                    Call Coll.Add(Block.GetBlock(0, -1))
'        Case TetrominoShape.ZShape: Call Coll.Add(Block.GetBlock(-1, -1))
'                                    Call Coll.Add(Block.GetBlock(-1, 0))
'                                    Call Coll.Add(Block.GetBlock(0, 0))
'                                    Call Coll.Add(Block.GetBlock(0, 1))
    End Select
    
    Set GetTetrominoBlocks = Coll
End Function
'-----------------------------------------------------------------------------80
Private Function GetTetrominoColor(ByVal Shape As TetrominoShape) As Long
Select Case Shape
    Case TetrominoShape.IShape: Let GetTetrominoColor = RGB(0, 255, 255)
    Case TetrominoShape.OShape: Let GetTetrominoColor = RGB(255, 255, 0)
    Case TetrominoShape.TShape: Let GetTetrominoColor = RGB(128, 0, 128)
    Case TetrominoShape.JShape: Let GetTetrominoColor = RGB(0, 0, 255)
    Case TetrominoShape.LShape: Let GetTetrominoColor = RGB(255, 128, 0)
    Case TetrominoShape.SShape: Let GetTetrominoColor = RGB(0, 255, 0)
    Case TetrominoShape.ZShape: Let GetTetrominoColor = RGB(255, 0, 0)
End Select
End Function
