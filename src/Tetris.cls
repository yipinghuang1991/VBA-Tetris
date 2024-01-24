VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tetris"
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
' Windows API
'=============================================================================80
'https://stackoverflow.com/questions/62937869/excel-vba-determine-mouse-location-when-clicking-on-a-cell
#If VBA7 Then
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If
'=============================================================================80
' Constants
'=============================================================================80
Const RowHeight As Double = 14.4
Const ColumnWidth As Long = 2
Const BackgroundColor As Long = 0
Const RepeatDelay As Long = 200
Const ReciprocalOfRepeatRate As Long = 35
'=============================================================================80
' Variable Container
'=============================================================================80
Private Type Variable
' Set Up
    ' -3: Error, -2: End Manually, -1: Topped Out, 0: Paused, 1: Ongoing, 2: Resumed
    GameState As Long
    HideBanner As Boolean
    Worksheet As Worksheet
    IsShowcase As Boolean
    Height As Long
    Width As Long
    Cell_1_1_Address As String
    Board As Range
    Matrix() As Cell
    Rows() As Variant
    FirstNonEmptyRow As Long
    ' Object
    ChangedCells As Dictionary
    BoundaryBlocks As Collection
    ' Next
    IncomingShapes As Collection
    Nexts() As Tetris
    ' Hold
    HeldShape As TetrominoShape
    Held As Tetris
    IsHeld As Boolean
    ' Piece
    FallingShape As TetrominoShape
    FallingPiece As Tetromino
    TestPiece As Tetromino
    ' Level
    Level As Long
    LinesCleared As Long
    LinesToLevelUp As Long
    Score As Long
    ' Time
    PieceStayTime As Long
    LastPressedKey As Integer
    LastMoveDownTime As Long
End Type
'-----------------------------------------------------------------------------80
Private This As Variable
'=============================================================================80
' Event
'=============================================================================80
Private Sub Class_Initialize()
    Call SpeedUp(True)
    Set This.ChangedCells = New Dictionary
End Sub
'-----------------------------------------------------------------------------80
Private Sub Class_Terminate()
    'Call This.Worksheet.Unprotect
    Call SpeedUp(False)
End Sub
'=============================================================================80
' Read-Only Property
'=============================================================================80
Public Property Get Self() As Tetris
    Set Self = Me
End Property
'=============================================================================80
' Wrire-Only Property
'=============================================================================80
Public Property Set Worksheet(ByRef Worksheet As Worksheet)
    Set This.Worksheet = Worksheet
End Property
'-----------------------------------------------------------------------------80
Public Property Let IsShowcase(ByVal IsShowcase As Boolean)
    Let This.IsShowcase = IsShowcase
End Property
'-----------------------------------------------------------------------------80
Public Property Let Height(ByVal Height As Long)
    Let This.Height = Height + IIf(This.IsShowcase, 0, 2)
End Property
'-----------------------------------------------------------------------------80
Public Property Let Width(ByVal Width As Long)
    Let This.Width = Width
End Property
'-----------------------------------------------------------------------------80
Public Property Let Cell_1_1_Address(ByVal Cell_1_1_Address As String)
With This
    Let .Cell_1_1_Address = Cell_1_1_Address
    Set .Board = .Worksheet.Range(.Cell_1_1_Address).Resize(.Height, .Width)
    Call InitializeCells
End With
End Property
'-----------------------------------------------------------------------------80
Private Property Let Level(ByVal Level As Long)
With This
    Let .Level = Level
    
    Dim LevelCap As Long: Let LevelCap = IIf(.Level >= 15, 15, .Level)
    Let .PieceStayTime = 1000 * (0.8 - 0.007 * (LevelCap - 1)) ^ (LevelCap - 1)
    Dim LinesToLevelUpPrev As Long: Let LinesToLevelUpPrev = IIf(.Level > 1, .LinesToLevelUp, 0)
    Let LinesToLevelUp = 10 * LevelCap
    Let LinesCleared = .LinesCleared - LinesToLevelUpPrev
    Let .Worksheet.Range("V15").Resize(, 3).Value2 = Array(.Level \ 100, _
                                                           .Level \ 10 Mod 10, _
                                                           .Level Mod 10)
End With
End Property
'-----------------------------------------------------------------------------80
Private Property Let LinesCleared(ByVal Lines As Long)
With This
    Let .LinesCleared = Lines
    Let .Worksheet.Range("U18").Resize(, 3).Value2 = Array(.LinesCleared \ 100, _
                                                           .LinesCleared \ 10 Mod 10, _
                                                           .LinesCleared Mod 10)
    If .LinesCleared >= .LinesToLevelUp Then Let Level = 1 + .Level
End With
End Property
'-----------------------------------------------------------------------------80
Private Property Let LinesToLevelUp(ByVal Lines As Long)
With This
    Let .LinesToLevelUp = Lines
    Let .Worksheet.Range("Y18").Resize(, 3).Value2 = Array(.LinesToLevelUp \ 100, _
                                                           .LinesToLevelUp \ 10 Mod 10, _
                                                           .LinesToLevelUp Mod 10)
End With
End Property
'-----------------------------------------------------------------------------80
Private Property Let Score(ByVal Score As Long)
With This
    Let .Score = Score
    Let .Worksheet.Range("U21").Resize(, 7).Value2 = Array(.Score \ 1000000, _
                                                           .Score \ 100000 Mod 10, _
                                                           .Score \ 10000 Mod 10, _
                                                           .Score \ 1000 Mod 10, _
                                                           .Score \ 100 Mod 10, _
                                                           .Score \ 10 Mod 10, _
                                                           .Score Mod 10)
End With
End Property
'-----------------------------------------------------------------------------80
Private Property Let Message(ByVal str As String)
With This
    Let .Worksheet.Range("U24").Value2 = str
End With
End Property
'=============================================================================80
' Read-Write Property
'=============================================================================80
Public Property Get GameState() As Long
    Let GameState = Self.GameState
End Property
'-----------------------------------------------------------------------------80
Public Property Let GameState(ByVal GameState As Long)
    Let This.GameState = GameState
    ' -3: Error, -2: End Manually, -1: Topped Out, 0: Paused, 1: Ongoing, 2: Resumed
    Select Case This.GameState
        Case -3: Let Message = "Error"
        Case -2: Let Message = "Ended"
        Case -1: Let Message = "Game Over"
        Case 0: Let Message = "Paused"
        Case 1: Let Message = "New Game"
        Case 2: Let Message = "Resumed"
    End Select
End Property
'=============================================================================80
' Constructor
'=============================================================================80
Public Function GetTetris(ByRef Worksheet As Worksheet, _
                          ByVal Cell_1_1_Address As String, _
                          Optional ByVal Height As Long = 20, _
                          Optional ByVal Width As Long = 10, _
                          Optional ByVal IsShowcase As Boolean) As Tetris
With New Tetris
    Set .Worksheet = Worksheet
    Let .IsShowcase = IsShowcase
    Let .Height = Height
    Let .Width = Width
    Let .Cell_1_1_Address = Cell_1_1_Address
    
    Set GetTetris = .Self
End With
End Function
'=============================================================================80
' Method
'=============================================================================80
Public Sub Start()
With This
    With ActiveWindow
        .DisplayHeadings = False
        .DisplayGridlines = False
    End With
    
    With .Worksheet
        Call .Unprotect
        Call .Range("A1").Select
        Let .Range("A1:AB31").RowHeight = RowHeight
        Let .Range("A1:AB31").ColumnWidth = ColumnWidth
        ' Make the adjacent row and column largest so as to draw less thing on the screen
        Let .Range("AC:AC").ColumnWidth = 255
        Let .Range("32:32").RowHeight = 409
        ' Make the adjacent row and column black
        Let .Range("AC1:AC32").Interior.ColorIndex = 1
        Let .Range("A32:AB32").Interior.ColorIndex = 1
        ' Hide the top two rows according to guideline
        'Let .Range("8:9").EntireRow.Hidden = True
        If This.HideBanner Then Let .Range("2:7").EntireRow.Hidden = True
        Call .Protect(, True, True, True, , True)
    End With
    
    If Not .IsShowcase Then
        Set .IncomingShapes = New Collection
        Call AddBag
        Set .Held = Tetris.GetTetris(.Worksheet, "U11", 2, 4, True)
        ReDim .Nexts(1 To 6) As Tetris
        
        Set .Nexts(1) = Tetris.GetTetris(.Worksheet, "C11", 2, 4, True)
        Set .Nexts(2) = Tetris.GetTetris(.Worksheet, "C14", 2, 4, True)
        Set .Nexts(3) = Tetris.GetTetris(.Worksheet, "C17", 2, 4, True)
        Set .Nexts(4) = Tetris.GetTetris(.Worksheet, "C20", 2, 4, True)
        Set .Nexts(5) = Tetris.GetTetris(.Worksheet, "C23", 2, 4, True)
        Set .Nexts(6) = Tetris.GetTetris(.Worksheet, "C26", 2, 4, True)
    End If
    
    ' Draw board, hold, and nexts
    Call DrawAll
    Call .Held.DrawAll
    
    Dim Tet As Variant
    For Each Tet In .Nexts
        Call Tet.DrawAll
    Next Tet
    
    ' Some settings
    Let Level = 1
    Let GameState = 1
    Let Score = 0
    Let .FirstNonEmptyRow = .Height
    Call SpawnIncoming(True, True)
    Call GameLoop
End With
End Sub
'-----------------------------------------------------------------------------80
Private Sub GameLoop()
    'On Error GoTo Err
    
    Dim Time As Long, LastRepeatTime As Long, LastKeyPressedTime As Long
    Dim LastMoveDownTime: Let LastMoveDownTime = TimerInMilliseconds()
    ' -2: End Manually, -1: Topped Out, 0: Paused, 1: Ongoing, 2: Resumed
    With This
   '| Decrease indentation level
    Do While .GameState >= 0
        If .GameState > 0 Then
            Let Time = TimerInMilliseconds()
            If Time - LastMoveDownTime >= .PieceStayTime Then
                Call Translate(1, 0, True, True)
                Let LastMoveDownTime = Time
            Else
                Select Case True 'KeyCodeConstants
               '| Decrease indentation level
                ' Space to hard drop.
                Case GetAsyncKeyState(vbKeySpace) And &H8000
                    If .LastPressedKey <> vbKeySpace Then
                        Let .LastPressedKey = vbKeySpace
                        Call HardDrop
                    End If
                ' Shift and C are to hold.
                Case GetAsyncKeyState(vbKeyShift) And &H8000
                    If .LastPressedKey <> vbKeyShift Then
                        Let .LastPressedKey = vbKeyShift
                        Call Hold
                    End If
                Case GetAsyncKeyState(vbKeyC) And &H8000
                    If .LastPressedKey <> vbKeyC Then
                        Let .LastPressedKey = vbKeyC
                        Call Hold
                    End If
                ' Up arrow and X are to rotate clockwise.
                Case GetAsyncKeyState(vbKeyUp) And &H8000
                    If .LastPressedKey <> vbKeyUp Then
                        Let .LastPressedKey = vbKeyUp
                        Call Rotate(-1, True, True)
                    End If
                Case GetAsyncKeyState(vbKeyX) And &H8000
                    If .LastPressedKey <> vbKeyX Then
                        Let .LastPressedKey = vbKeyX
                        Call Rotate(-1, True, True)
                    End If
                ' Ctrl and Z are to rotate counterclockwise.
                Case GetAsyncKeyState(vbKeyControl) And &H8000
                    If .LastPressedKey <> vbKeyControl Then
                        Let .LastPressedKey = vbKeyControl
                        Call Rotate(1, True, True)
                    End If
                Case GetAsyncKeyState(vbKeyZ) And &H8000
                    If .LastPressedKey <> vbKeyZ Then
                        Let .LastPressedKey = vbKeyZ
                        Call Rotate(1, True, True)
                    End If
                ' Left, right, and down arrows are the same as on the console.
                Case GetAsyncKeyState(vbKeyLeft) And &H8000
                    If .LastPressedKey <> vbKeyLeft Then
                        Let LastKeyPressedTime = TimerInMilliseconds()
                        Let .LastPressedKey = vbKeyLeft
                        Call Translate(0, -1, True, True)
                    Else
                        Let Time = TimerInMilliseconds()
                        If Time - LastKeyPressedTime >= RepeatDelay Then
                            If Time - LastRepeatTime >= ReciprocalOfRepeatRate Then
                                Let LastRepeatTime = Time
                                Call Translate(0, -1, True, True)
                            End If
                        End If
                    End If
                Case GetAsyncKeyState(vbKeyRight) And &H8000
                    If .LastPressedKey <> vbKeyRight Then
                        Let LastKeyPressedTime = TimerInMilliseconds()
                        Let .LastPressedKey = vbKeyRight
                        Call Translate(0, 1, True, True)
                    Else
                        Let Time = TimerInMilliseconds()
                        If Time - LastKeyPressedTime >= RepeatDelay Then
                            If Time - LastRepeatTime >= ReciprocalOfRepeatRate Then
                                Let LastRepeatTime = Time
                                Call Translate(0, 1, True, True)
                            End If
                        End If
                    End If
                Case GetAsyncKeyState(vbKeyDown) And &H8000
                    If .LastPressedKey <> vbKeyDown Then
                        Let LastKeyPressedTime = TimerInMilliseconds()
                        Let .LastPressedKey = vbKeyDown
                        Call SoftDown
                    Else
                        Let Time = TimerInMilliseconds()
                        If Time - LastKeyPressedTime >= RepeatDelay Then
                            If Time - LastRepeatTime >= ReciprocalOfRepeatRate Then
                                Let LastRepeatTime = Time
                                Call SoftDown
                            End If
                        End If
                    End If
                ' Pause
                Case GetAsyncKeyState(vbKeyP) And &H8000
                    Let GameState = 0
                ' End
                Case GetAsyncKeyState(vbKeyEscape) And &H8000
                    Let GameState = -2
                Case Else
                    If .LastPressedKey <> 0 Then Let .LastPressedKey = 0
               '| Decrease indentation level
                End Select
            End If
        Else ' .GameState = 0
            If GetAsyncKeyState(vbKeyR) And &H8000 Then _
                Let GameState = 2
        End If
        
        DoEvents
    Loop
   '| Decrease indentation level
    End With
    
    Exit Sub
Err:
    Let GameState = -3
End Sub
'-----------------------------------------------------------------------------80
Private Sub InitializeCells()
With This
    Dim X As Long, Y As Long
    
    ReDim .Matrix(1 To .Height, 1 To .Width) As Cell
    ReDim .Rows(1 To .Height) As Variant
    
    Dim Arr() As Cell
    
    For X = 1 To .Height
        ReDim Arr(1 To .Width) As Cell
        Let .Rows(X) = Arr
    Next X
    
    For X = 1 To .Height
        For Y = 1 To .Width
            Set .Matrix(X, Y) = New Cell
            Let .Matrix(X, Y).BackgroundColor = BackgroundColor
            Set .Matrix(X, Y).Range = .Board(X, Y)
        Next Y
    Next X
    
    For X = 1 To .Height
        For Y = 1 To .Width
            Set .Rows(X)(Y) = .Matrix(X, Y)
        Next Y
    Next X
End With
End Sub
'-----------------------------------------------------------------------------80
Public Sub DrawAll()
With This
    With .Board
        Let .Borders(XlBordersIndex.xlInsideHorizontal).Weight = XlBorderWeight.xlMedium
        Let .Borders(XlBordersIndex.xlInsideHorizontal).Color = RGB(31, 31, 31)
        Let .Borders(XlBordersIndex.xlInsideVertical).Weight = XlBorderWeight.xlMedium
        Let .Borders(XlBordersIndex.xlInsideVertical).Color = RGB(31, 31, 31)
        Call .BorderAround(, XlBorderWeight.xlMedium)
    End With
    
    Dim C As Variant
    For Each C In .Matrix
        Call C.Draw
    Next C
End With
End Sub
'-----------------------------------------------------------------------------80
Private Sub AddChangedCell(ByRef ChangedCell As Variant)
    On Error Resume Next
    Call This.ChangedCells.Add(ChangedCell, vbNullString)
End Sub
'-----------------------------------------------------------------------------80
Public Sub DrawChanged()
With This
    'Let Application.ScreenUpdating = False
    
    If .ChangedCells.Count Then
        Dim C As Variant
        For Each C In .ChangedCells.Keys
            Call C.Draw
        Next C
        
        Set .ChangedCells = New Dictionary
    End If
    
    'Let Application.ScreenUpdating = True
End With
End Sub
'=============================================================================80
' Method: Tetromino
'=============================================================================80
Private Sub LinkPiece()
With This
    Dim B As Block
    For Each B In .FallingPiece.Blocks
        Set .Matrix(B.X, B.Y).Block = B
        Call AddChangedCell(.Matrix(B.X, B.Y))
    Next B
End With
End Sub
'-----------------------------------------------------------------------------80
Private Sub UnlinkPiece()
With This
    If Not .FallingPiece Is Nothing Then
        Dim B As Block
        For Each B In .FallingPiece.Blocks
            Set .Matrix(B.X, B.Y).Block = Nothing
            Call AddChangedCell(.Matrix(B.X, B.Y))
        Next B
    End If
End With
End Sub
'-----------------------------------------------------------------------------80
Private Sub AcceptTestPiece(Optional ByVal Draw As Boolean)
With This
    Call UnlinkPiece
    Set .FallingPiece = .TestPiece
    Call LinkPiece
    If Draw Then Call DrawChanged
End With
End Sub
'-----------------------------------------------------------------------------80
Private Function IsTestPieceOverlapped() As Boolean
With This
    Dim B As Block
    For Each B In .TestPiece.Blocks
        If B.X > .Height Or B.Y < 1 Or B.Y > .Width Then
            Let IsTestPieceOverlapped = True
            Exit For
        ElseIf .Matrix(B.X, B.Y).State = 0 Then
            Let IsTestPieceOverlapped = True
            Exit For
        End If
    Next B
End With
End Function
'-----------------------------------------------------------------------------80
Private Sub AddBag()
    Dim v As Variant
    For Each v In ShuffleArray(Array(1, 2, 3, 4, 5, 6, 7))
        Call This.IncomingShapes.Add(v)
    Next v
End Sub
'-----------------------------------------------------------------------------80
Private Function SpawnIncoming(Optional ByVal AcceptOnSuccess As Boolean, _
                               Optional ByVal DrawOnSuccess As Boolean) As Boolean
With This
    Set .TestPiece = Tetromino.GetTetromino(.IncomingShapes(1), IIf(.IncomingShapes(1) = TetrominoShape.OShape, 1, 2), IIf(.IsShowcase, 2, 5))
    
    If Not IsTestPieceOverlapped() Then
        Let SpawnIncoming = True
        If AcceptOnSuccess Then
            Let .FallingShape = .IncomingShapes(1)
            Call .IncomingShapes.Remove(1)
            If .IncomingShapes.Count = 6 Then Call AddBag
        
            Dim k As Long
            For k = 1 To 6
                Call .Nexts(k).Spawn(.IncomingShapes(k))
            Next k
            
            Call Spawn
            'Call AcceptTestPiece(DrawOnSuccess)
        End If
    Else
        Let GameState = -1
    End If
End With
End Function
'-----------------------------------------------------------------------------80
Public Sub Spawn(Optional ByVal Shape As TetrominoShape)
With This
    If .IsShowcase And (Not .FallingPiece Is Nothing) Then Call UnlinkPiece
    If Shape <> 0 Then Let .FallingShape = Shape
    Set .FallingPiece = Tetromino.GetTetromino(.FallingShape, IIf(.FallingShape = TetrominoShape.OShape, 1, 2), IIf(.IsShowcase, 2, 5))
    Call LinkPiece
    Call DrawChanged
End With
End Sub
'-----------------------------------------------------------------------------80
Private Function Hold() As Boolean
With This
    If .IsHeld Then
        Let Message = "Already held"
    Else
        Let .IsHeld = True
        
        Call UnlinkPiece
        
        Dim TempShape As TetrominoShape: Let TempShape = .HeldShape
        Let .HeldShape = .FallingShape
        
        If TempShape <> 0 Then
            Let .FallingShape = TempShape
            Call Spawn
        Else
            Call SpawnIncoming(True, True)
        End If
        
        Let Message = "Hold"
        Call .Held.Spawn(.HeldShape)
        
        Let Hold = True
    End If
End With
End Function
'=============================================================================80
' Method: Translation
'=============================================================================80
Private Function Translate(ByVal X As Long, ByVal Y As Long, _
                           Optional ByVal AcceptOnSuccess As Boolean, _
                           Optional ByVal DrawOnSuccess As Boolean) As Boolean
With This
    Set .TestPiece = .FallingPiece.Copy
    Call .TestPiece.Translate(X, Y)
    If Not IsTestPieceOverlapped() Then
        Let Translate = True
        If AcceptOnSuccess Then Call AcceptTestPiece(DrawOnSuccess)
    ElseIf X = 1 Then
        Call Ground
    End If
End With
End Function
'-----------------------------------------------------------------------------80
Private Sub SoftDown()
    If Translate(1, 0, True, True) Then Call ScoreAdd(1)
End Sub
'-----------------------------------------------------------------------------80
Private Sub HardDrop()
With This
    Dim X As Long
    Set .TestPiece = .FallingPiece.Copy
    
    Do
        Let X = 1 + X
        Call .TestPiece.Translate(1, 0)
    Loop While Not IsTestPieceOverlapped()
    
    Call .TestPiece.Translate(-1, 0)
    Call AcceptTestPiece(True)
    Call Ground
    Call ScoreAdd(2 * (X - 1))
End With
End Sub
'-----------------------------------------------------------------------------80
Private Sub Ground()
With This
    Let .IsHeld = False
    
    Dim B As Block
    For Each B In .FallingPiece.Blocks
        Set B.Parent = Nothing
        Set .Matrix(B.X, B.Y).Block = B
        If B.X < .FirstNonEmptyRow Then Let .FirstNonEmptyRow = B.X
    Next B
    
    Let Message = vbNullString
    Call ClearLines
    Call SpawnIncoming(True, True)
End With
End Sub
'=============================================================================80
' Method: Rotation
'=============================================================================80
' Stance of piece if rotate CW (or CCW) N times
' | N   | 0 | 1 | 2 | 3 |
' |-----|---|---|---|---|
' | CCW | 0 | 1 | 2 | 3 |
' | CW  | 0 | 3 | 2 | 1 |
' If basic rotation fails, from the rotated stance, test by moving piece
'
' SRS Kick Table for shape T, J, L, S, Z:
' ---------------------------------------------------
' |From| To |  Test1  |  Test2  |  Test3  |  Test4  | Transition
' |----|----|---------|---------|---------|---------|
' |  0 |  1 | ( 0,-1) | ( 1,-1) | (-2, 0) | (-2,-1) |   1
' |  1 |  2 | ( 0, 1) | (-1, 1) | ( 2, 0) | ( 2, 1) |  12
' |  2 |  3 | ( 0, 1) | ( 1, 1) | (-2, 0) | (-2, 1) |  23
' |  3 |  0 | ( 0,-1) | (-1,-1) | ( 2, 0) | ( 2,-1) |  30
' |  0 |  3 |     Negetive of transition 3 -> 0     |   3
' |  3 |  2 |     Negetive of transition 2 -> 3     |  32
' |  2 |  1 |     Negetive of transition 1 -> 2     |  21
' |  1 |  0 |     Negetive of transition 0 -> 1     |  10
' ---------------------------------------------------
'
' SRS Kick Table for shape I:
' ---------------------------------------------------
' |From| To |  Test1  |  Test2  |  Test3  |  Test4  | Transition
' |----|----|---------|---------|---------|---------|
' |  0 |  1 | ( 0, 1) | ( 0,-2) | ( 2, 1) | (-1,-2) |   1
' |  1 |  2 | ( 0, 2) | ( 0,-1) | (-1, 2) | ( 2,-1) |  12
' |  2 |  3 |     Negetive of transition 0 -> 1     |  23
' |  3 |  0 |     Negetive of transition 1 -> 2     |  30
' |  0 |  3 |     Negetive of transition 3 -> 0     |   3
' |  3 |  2 |     Negetive of transition 2 -> 3     |  32
' |  2 |  1 |     Negetive of transition 1 -> 2     |  21
' |  1 |  0 |     Negetive of transition 0 -> 1     |  10
' ---------------------------------------------------
' (0 -> 3) = (1 -> 2)
' (3 -> 2) = (0 -> 1)
' (2 -> 1) = (3 -> 0) = - (1 -> 2)
' (1 -> 0) = (2 -> 3) = - (0 -> 1)
'-----------------------------------------------------------------------------80
Private Function Rotate(ByVal Sign As Long, _
                        Optional ByVal AcceptOnSuccess As Boolean, _
                        Optional ByVal DrawOnSuccess As Boolean) As Boolean
With This
    Set .TestPiece = .FallingPiece.Copy
    Call .TestPiece.Rotate(Sign) ' CCW: Sign = 1, CW: Sign = -1
    If Not IsTestPieceOverlapped() Then
        Let Rotate = True
        If AcceptOnSuccess Then Call AcceptTestPiece(DrawOnSuccess)
'    ElseIf Kick() Then
'        Let Rotate = True
'        If AcceptOnSuccess Then Call AcceptTestPiece(DrawOnSuccess)
    End If
End With
End Function
'-----------------------------------------------------------------------------80
Private Function Kick() As Boolean
With This
    Dim Transition As Long: Let Transition = 10 * .FallingPiece.Stance + .TestPiece.Stance
    Dim Tests As Variant, Sign As Long
    
    If .FallingShape = TetrominoShape.IShape Then
        Select Case Transition
            Case 1, 10, 23, 32: Let Tests = Array(Array(0, 1), Array(0, -2), Array(2, 1), Array(-1, -2))
            Case 3, 12, 21, 30: Let Tests = Array(Array(0, 2), Array(0, -1), Array(-1, 2), Array(2, -1))
        End Select
        
        Select Case Transition
            Case 1, 3, 12, 32: Let Sign = 1
            Case Else: Let Sign = -1
        End Select
    Else
        Select Case Transition
            Case 1, 10: Let Tests = Array(Array(0, -1), Array(1, -1), Array(-2, 0), Array(-2, -1))
            Case 12, 21: Let Tests = Array(Array(0, 1), Array(-1, 1), Array(2, 0), Array(2, 1))
            Case 23, 32: Let Tests = Array(Array(0, 1), Array(1, 1), Array(-2, 0), Array(-2, 1))
            Case 3, 30: Let Tests = Array(Array(0, -1), Array(-1, -1), Array(2, 0), Array(2, -1))
        End Select
        
        Select Case Transition
            Case 1, 12, 23, 30: Let Sign = 1
            Case Else: Let Sign = -1
        End Select
    End If
    
    Dim test As Variant
    For Each test In Tests
        Call .TestPiece.Translate(Sign * test(0), Sign * test(1))
        Debug.Print "Center: X: " & .TestPiece.X & ", Y: " & .TestPiece.Y
        Debug.Print "Transition: " & Transition & ", X: " & Sign * test(0) & ", Y: " & Sign * test(1)
        
        If Not IsTestPieceOverlapped() Then
            Let Kick = True
            Exit For
        End If
    Next test
End With
End Function
'=============================================================================80
' Method: Clearing
'=============================================================================80
Private Sub ClearLines()
With This
    If .FirstNonEmptyRow = .Height Then Exit Sub
    
    Dim X As Long, IsLine As Boolean, C As Variant, Xv As Variant
    
    Dim Lines As Collection: Set Lines = New Collection
    Dim RowMoveDownStep() As Long: ReDim RowMoveDownStep(1 To .Height) As Long
    'Dim MoveRow() As Long: ReDim MoveRow(.FirstNonEmptyRow To .Height - 1) As Long
    
'    For X = .FirstNonEmptyRow To .Height - 1
'        Let MoveRow(X) = X
'    Next X
    
    ' Determine row number of lines
    For X = .Height To .FirstNonEmptyRow Step -1
        Let IsLine = True
        
        For Each C In .Rows(X)
            If C.State <> 0 Then
                Let IsLine = False
                Exit For
            End If
        Next C
        
        If IsLine Then Call Lines.Add(X)
    Next X
    
    If Lines.Count Then
        For Each Xv In Lines
            For X = .FirstNonEmptyRow To Xv - 1
                Let RowMoveDownStep(X) = 1 + RowMoveDownStep(X)
                'Let MoveRow(X) = 1 + MoveRow(X)
            Next X
            'Let MoveRow(X) = 0
            Let RowMoveDownStep(Xv) = 0
            ' Delete Row
            For Each C In .Rows(Xv)
                Set C.Block = Nothing
                Call AddChangedCell(C)
            Next C
        Next Xv
        
        For X = Lines(1) - 1 To .FirstNonEmptyRow Step -1
            If RowMoveDownStep(X) <> 0 Then Call MoveRowDown(X, RowMoveDownStep(X))
            'If MoveRow(X) <> 0 Then Call MoveRowToRow(X, MoveRow(X))
        Next X
        
        Let .FirstNonEmptyRow = Lines.Count + .FirstNonEmptyRow
        Let LinesCleared = Lines.Count + .LinesCleared
        Call ScoreAdd(100 * Choose(Lines.Count, 1, 3, 5, 8) * .Level)
        Let Message = Choose(Lines.Count, "Single", "Double", "Triple", "Tetris!")
    End If
End With
End Sub
'-----------------------------------------------------------------------------80
'Private Sub MoveRowToRow(ByVal X As Long, ByVal NewX As Long)
'With This
'    Dim C As Variant
'    For Each C In .Rows(X)
'        If C.State = 0 Then
'            With .Matrix(Step + C.Block.X, C.Block.Y)
'                Set .Block = C.Block
'                Let .Block.X = Step + .Block.X
'                Set C.Block = Nothing
'                Call AddChangedCell(.Self)
'            End With
'
'            Call AddChangedCell(C)
'        End If
'    Next C
'End With
'End Sub
Private Sub MoveRowDown(ByVal X As Long, ByVal Step As Long)
With This
    Dim C As Variant
    For Each C In .Rows(X)
        If C.State = 0 Then
            With .Matrix(Step + C.Block.X, C.Block.Y)
                Set .Block = C.Block
                Let .Block.X = Step + .Block.X
                Set C.Block = Nothing
                Call AddChangedCell(.Self)
            End With
            
            Call AddChangedCell(C)
        End If
    Next C
End With
End Sub
'=============================================================================80
' Helper Function
'=============================================================================80
Private Sub ScoreAdd(ByVal ScoreToAdd As Long)
    Let Score = ScoreToAdd + This.Score
End Sub
'-----------------------------------------------------------------------------80
Private Function TimerInMilliseconds() As Long
    Let TimerInMilliseconds = 1000 * Timer
End Function
'-----------------------------------------------------------------------------80
' http://www.cpearson.com/excel/ShuffleArray.aspx
Private Function ShuffleArray(InArray As Variant) As Variant
    Dim N As Long, k As Long, L As Long
    Dim temp As Long, Arr() As Variant
    
    Call Randomize
    Let L = UBound(InArray) - LBound(InArray) + 1
    ReDim Arr(LBound(InArray) To UBound(InArray))
    For N = LBound(InArray) To UBound(InArray)
        Let Arr(N) = InArray(N)
    Next N
    
    For N = LBound(InArray) To UBound(InArray)
        Let k = CLng(((UBound(InArray) - N) * Rnd) + N)
        Let temp = Arr(N)
        Let Arr(N) = Arr(k)
        Let Arr(k) = temp
    Next N
    
    Let ShuffleArray = Arr
End Function
'-----------------------------------------------------------------------------80
Private Sub SpeedUp(ByVal Toggle As Boolean, Optional ByVal Calc As Variant)
With Excel.Application
    Let .AskToUpdateLinks = Not Toggle
    Let .Calculation = IIf(Toggle, xlCalculationManual, _
                       IIf(IsMissing(Calc), xlCalculationAutomatic, Calc))
    Let .DisplayAlerts = Not Toggle
    Let .DisplayFormulaBar = Not Toggle
    Let .DisplayScrollBars = Not Toggle
    Let .DisplayStatusBar = Not Toggle
    Let .EnableAnimations = Not Toggle
    Let .EnableEvents = Not Toggle
    Let .PrintCommunication = Not Toggle
    'Let .ScreenUpdating = Not toggle
    Call .ExecuteExcel4Macro("Show.Toolbar(""Ribbon""," & CStr(Not Toggle) & ")")
    
    With .ActiveWindow
        Let .DisplayGridlines = Not Toggle
        Let .DisplayHeadings = Not Toggle
        'Let .DisplayWorkbookTabs = Not Toggle
    End With
End With
End Sub
