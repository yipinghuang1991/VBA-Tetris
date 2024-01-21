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
'=============================================================================80
' Enumeration
'=============================================================================80
Public Enum TetrominoShape
    I = 1
    O
    T
    J
    L
    s
    z
End Enum
'=============================================================================80
' Variable Container
'=============================================================================80
Private Type Variable
    X As Long
    Y As Long
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
    
    If Shape = TetrominoShape.L Or Shape = TetrominoShape.z Then
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
Public Sub Move(ByVal XAxis As Long, ByVal YAxis As Long)
With This
    Let .X = XAxis + .X
    Let .Y = YAxis + .Y
End With
End Sub
'-----------------------------------------------------------------------------80
Public Sub Rotate(Optional ByVal Sign As Long = -1)
    Dim B As Block
    For Each B In This.Blocks
        Call B.Rotate(Sign)
    Next B
End Sub
'-----------------------------------------------------------------------------80
Public Sub Reflect(Optional ByVal XAxis As Boolean = True, Optional ByVal YAxis As Boolean)
    Dim B As Block
    For Each B In This.Blocks
        Call B.Reflect(XAxis, YAxis)
    Next B
End Sub
'=============================================================================80
' Helper Function
'=============================================================================80
Private Function GetTetrominoBlocks(ByVal Shape As TetrominoShape) As Collection
    Dim Coll As Collection: Set Coll = New Collection
    
    Select Case Shape
        Case TetrominoShape.I: Call Coll.Add(Block.GetBlock(0, -1, True))
                               Call Coll.Add(Block.GetBlock(0, 0, True))
                               Call Coll.Add(Block.GetBlock(0, 1, True))
                               Call Coll.Add(Block.GetBlock(0, 2, True))
        Case TetrominoShape.O: Call Coll.Add(Block.GetBlock(0, 0, True))
                               Call Coll.Add(Block.GetBlock(0, 1, True))
                               Call Coll.Add(Block.GetBlock(1, 1, True))
                               Call Coll.Add(Block.GetBlock(1, 0, True))
        Case TetrominoShape.T: Call Coll.Add(Block.GetBlock(-1, 0))
                               Call Coll.Add(Block.GetBlock(0, -1))
                               Call Coll.Add(Block.GetBlock(0, 0))
                               Call Coll.Add(Block.GetBlock(0, 1))
        Case TetrominoShape.J, TetrominoShape.L
                               Call Coll.Add(Block.GetBlock(0, 1))
                               Call Coll.Add(Block.GetBlock(0, 0))
                               Call Coll.Add(Block.GetBlock(0, -1))
                               Call Coll.Add(Block.GetBlock(-1, -1))
'        Case TetrominoShape.L: Call Coll.Add(Block.GetBlock(0, -1))
'                               Call Coll.Add(Block.GetBlock(0, 0))
'                               Call Coll.Add(Block.GetBlock(0, 1))
'                               Call Coll.Add(Block.GetBlock(-1, 1))
        Case TetrominoShape.s, TetrominoShape.z
                               Call Coll.Add(Block.GetBlock(-1, 1))
                               Call Coll.Add(Block.GetBlock(-1, 0))
                               Call Coll.Add(Block.GetBlock(0, 0))
                               Call Coll.Add(Block.GetBlock(0, -1))
'        Case TetrominoShape.Z: Call Coll.Add(Block.GetBlock(-1, -1))
'                               Call Coll.Add(Block.GetBlock(-1, 0))
'                               Call Coll.Add(Block.GetBlock(0, 0))
'                               Call Coll.Add(Block.GetBlock(0, 1))
    End Select
    
    Set GetTetrominoBlocks = Coll
End Function
'-----------------------------------------------------------------------------80
Private Function GetTetrominoColor(ByVal Shape As TetrominoShape) As Long
Select Case Shape
    Case TetrominoShape.I: Let GetTetrominoColor = RGB(0, 255, 255)
    Case TetrominoShape.O: Let GetTetrominoColor = RGB(255, 255, 0)
    Case TetrominoShape.T: Let GetTetrominoColor = RGB(128, 0, 128)
    Case TetrominoShape.J: Let GetTetrominoColor = RGB(0, 0, 255)
    Case TetrominoShape.L: Let GetTetrominoColor = RGB(255, 128, 0)
    Case TetrominoShape.s: Let GetTetrominoColor = RGB(0, 255, 0)
    Case TetrominoShape.z: Let GetTetrominoColor = RGB(255, 0, 0)
End Select
End Function
