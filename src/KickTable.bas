Attribute VB_Name = "KickTable"
Option Explicit
'=============================================================================80
' Kick Table
'
' Author: Yi-Ping Huang
' https://github.com/yipinghuang1991/VBA-Tetris
'=============================================================================80
' Stance of piece if rotate CW (or CCW) N times
' | N   | 0 | 1 | 2 | 3 |
' |-----|---|---|---|---|
' | CCW | 0 | 1 | 2 | 3 |
' | CW  | 0 | 3 | 2 | 1 |
' If basic rotation fails, from the rotated stance, test by moving piece
'-----------------------------------------------------------------------------80
' SRS Kick Table for shape I:
' -------------------------------------------------------------------------------------------------
' |From| To |  Test1  |  Test2  |  Test3  |  Test4  |Trans|  Test1  |  T2-T1  |  T3-T2  |  T4-T3  |
' |----|----|---------|---------|---------|---------|-----|---------|---------|---------|---------|
' |  0 |  1 | ( 0,-1) | ( 0, 2) | (-2,-1) | ( 1, 2) |   1 | ( 0,-1) | ( 0, 3) | (-2,-3) | ( 3, 3) |
' |  1 |  2 | ( 0,-2) | ( 0, 1) | ( 1,-2) | (-2, 1) |  12 | ( 0,-2) | ( 0, 3) | ( 1,-3) | (-3, 3) |
' |  2 |  3 |     Negetive of transition 0 -> 1     |  23 |                                       |
' |  3 |  0 |     Negetive of transition 1 -> 2     |  30 |                                       |
' |  0 |  3 |     Negetive of transition 3 -> 0     |   3 |                                       |
' |  3 |  2 |     Negetive of transition 2 -> 3     |  32 |                                       |
' |  2 |  1 |     Negetive of transition 1 -> 2     |  21 |                                       |
' |  1 |  0 |     Negetive of transition 0 -> 1     |  10 |                                       |
' -------------------------------------------------------------------------------------------------
' (0 -> 3) = (1 -> 2)
' (3 -> 2) = (0 -> 1)
' (2 -> 1) = (3 -> 0) = - (1 -> 2)
' (1 -> 0) = (2 -> 3) = - (0 -> 1)
'-----------------------------------------------------------------------------80
' SRS Kick Table for shape T, J, L, S, Z:
' -------------------------------------------------------------------------------------------------
' |From| To |  Test1  |  Test2  |  Test3  |  Test4  |Trans|  Test1  |  T2-T1  |  T3-T2  |  T4-T3  |
' |----|----|---------|---------|---------|---------|-----|---------|---------|---------|---------|
' |  0 |  1 | ( 0, 1) | (-1, 1) | ( 2, 0) | ( 2, 1) |   1 | ( 0, 1) | (-1, 0) | ( 3,-1) | ( 0, 1) |
' |  1 |  2 |     Negetive of transition 0 -> 1     |  12 |                                       |
' |  2 |  3 | ( 0,-1) | (-1,-1) | ( 2, 0) | ( 2,-1) |  23 | ( 0,-1) | (-1, 0) | ( 3, 1) | ( 0,-1) |
' |  3 |  0 |     Negetive of transition 2 -> 3     |  30 |                                       |
' |  0 |  3 |     Negetive of transition 3 -> 0     |   3 |                                       |
' |  3 |  2 |     Negetive of transition 2 -> 3     |  32 |                                       |
' |  2 |  1 |     Negetive of transition 1 -> 2     |  21 |                                       |
' |  1 |  0 |     Negetive of transition 0 -> 1     |  10 |                                       |
' -------------------------------------------------------------------------------------------------
' (0 -> 3) = (2 -> 3)
' (3 -> 2) = (3 -> 0) = - (2 -> 3)
' (2 -> 1) = (0 -> 1)
' (1 -> 0) = (1 -> 2) = - (0 -> 1)
'-----------------------------------------------------------------------------80
Public Function GetKickTests(ByVal Shape As TetrominoShape, _
                             ByVal FromStance As Long, ByVal ToStance As Long) As Variant
    Dim Tests As Variant, Sign As Long
    Dim Transition As Long: Let Transition = 10 * FromStance + ToStance
    
    Select Case Shape
        Case IShape
            Select Case Transition
                Case 1, 10, 23, 32: Let Tests = Array(Array(0, -1), Array(0, 3), Array(-2, -3), Array(3, 3))
                Case 3, 12, 21, 30: Let Tests = Array(Array(0, -2), Array(0, 3), Array(1, -3), Array(-3, 3))
            End Select
            
            Select Case Transition
                Case 1, 3, 12, 32: Let Sign = 1
                Case Else: Let Sign = -1
            End Select
        Case TShape, JShape, LShape, SShape, ZShape
            Select Case Transition
                Case 1, 10, 12, 21: Let Tests = Array(Array(0, 1), Array(-1, 0), Array(3, -1), Array(0, 1))
                Case 3, 23, 30, 32: Let Tests = Array(Array(0, -1), Array(-1, 0), Array(3, 1), Array(0, -1))
            End Select
            
            Select Case Transition
                Case 1, 3, 21, 23: Let Sign = 1
                Case Else: Let Sign = -1
            End Select
    End Select
    
    If IsArray(Tests) Then
        Let Tests(0)(0) = Sign * Tests(0)(0)
        Let Tests(0)(1) = Sign * Tests(0)(1)
        Let Tests(1)(0) = Sign * Tests(1)(0)
        Let Tests(1)(1) = Sign * Tests(1)(1)
        Let Tests(2)(0) = Sign * Tests(2)(0)
        Let Tests(2)(1) = Sign * Tests(2)(1)
        Let Tests(3)(0) = Sign * Tests(3)(0)
        Let Tests(3)(1) = Sign * Tests(3)(1)
        
        Let GetKickTests = Tests
    End If
End Function


