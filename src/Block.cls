VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Block"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'=============================================================================80
' A block is either
' (1) The building blocks of Tetrominos
'     In this case, the (x, y) is relative to the center of the Tetromino
' (2) The grounded blocks
'     In this case, the (x, y) is the position on the board
'
' Author: Yi-Ping Huang
' https://github.com/yipinghuang1991/VBA-Tetris
'=============================================================================80
'=============================================================================80
' Variable Container
'=============================================================================80
Private Type Variable
    HasParent As Boolean
    Parent As Tetromino
    IsZeroCollapsed As Boolean
    x As Long
    y As Long
    Color As Long
End Type
'-----------------------------------------------------------------------------80
Private This As Variable
'=============================================================================80
' Read-Only Property
'=============================================================================80
Public Property Get Self() As Block
    Set Self = Me
End Property
'-----------------------------------------------------------------------------80
Public Property Get HasParent() As Boolean
    Let HasParent = This.HasParent
End Property
'-----------------------------------------------------------------------------80
Public Property Set Parent(ByRef Parent As Tetromino)
With This
    If .HasParent And (Parent Is Nothing) Then
        Let .HasParent = False
        Let .x = .Parent.x + .x
        Let .y = .Parent.y + .y
        Let .Color = .Parent.Color
    ElseIf (Not .HasParent) And (Not Parent Is Nothing) Then
        Let .HasParent = True
    End If
    
    Set .Parent = Parent
End With
End Property
'=============================================================================80
' Write-Only Property
'=============================================================================80
Public Property Let IsZeroCollapsed(ByVal IsZeroCollapsed As Boolean)
    Let This.IsZeroCollapsed = IsZeroCollapsed
End Property
'=============================================================================80
' Read-Write Property
'=============================================================================80
Private Property Get x() As Long
With This
    If .HasParent Then
        Let x = .Parent.x + .x
    Else
        Let x = .x
    End If
End With
End Property
Public Property Let x(ByVal x As Long)
    Let This.x = x
End Property
'-----------------------------------------------------------------------------80
Private Property Get y() As Long
With This
    If .HasParent Then
        Let y = .Parent.y + .y
    Else
        Let y = .y
    End If
End With
End Property
Public Property Let y(ByVal y As Long)
    Let This.y = y
End Property
'-----------------------------------------------------------------------------80
Public Property Get Color() As Long
With This
    If .HasParent Then
        Let Color = .Parent.Color
    Else
        Let Color = .Color
    End If
End With
End Property
Public Property Let Color(ByVal Color As Long)
    Let This.Color = Color
End Property
'=============================================================================80
' Constructor
'=============================================================================80
Public Function GetBlock(ByVal x As Long, ByVal y As Long, _
                         Optional ByVal IsZeroCollapsed As Boolean, _
                         Optional ByRef Parent As Tetromino, _
                         Optional ByVal Color As Long = 16777215) As Block
With New Block
    Let .IsZeroCollapsed = IsZeroCollapsed
    Let .x = x
    Let .y = y
    
    If Parent Is Nothing Then
        Let .Color = Color
    Else
        Set .Parent = Parent
    End If
    
    Set GetBlock = .Self
End With
End Function
'-----------------------------------------------------------------------------80
Public Function Copy() As Block
With This
    Set Copy = GetBlock(.x, .y, .IsZeroCollapsed, .Parent, .Color)
End With
End Function
'=============================================================================80
' Method
'=============================================================================80
Public Sub Rotate(Optional ByVal Sign As Long = -1)
With This
    If .IsZeroCollapsed Then
        Call UncollapseZero
        Call Rotate(Sign)
        Call CollapsedZero
    Else
        Dim XSave As Long: Let XSave = .x
        Let .x = -Sign * .y
        Let .y = Sign * XSave
    End If
End With
End Sub
'-----------------------------------------------------------------------------80
Public Sub Reflect(Optional ByVal XAxis As Boolean = True, Optional ByVal YAxis As Boolean)
With This
    If .IsZeroCollapsed Then
        Call UncollapseZero
        Call Reflect(XAxis, YAxis)
        Call CollapsedZero
    Else
        If XAxis Then Let .y = -.y
        If YAxis Then Let .x = -.x
    End If
End With
End Sub
'=============================================================================80
' Helper Function
'=============================================================================80
Private Sub CollapsedZero()
With This
    If .x < 0 Then Let .x = .x + 1
    If .y < 0 Then Let .y = .y + 1
    Let .IsZeroCollapsed = True
End With
End Sub
'-----------------------------------------------------------------------------80
Private Sub UncollapseZero()
With This
    If .x <= 0 Then Let .x = .x - 1
    If .y <= 0 Then Let .y = .y - 1
    Let .IsZeroCollapsed = False
End With
End Sub
