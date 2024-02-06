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
'
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
' Enumeration
'=============================================================================80
Public Enum GameState
    Error = -9
    EndManually
    ToppedOut
    Paused = 0
    Started
    Resumed
End Enum
'=============================================================================80
' Variable Container
'=============================================================================80
Private Type Variable
    State As GameState
    Worksheet As Worksheet
    IsShowcase As Boolean
    Height As Long
    Width As Long
    Cell_1_1_Address As String
    Board As Range
    Matrix() As Cell
    FirstNonEmptyRow As Long
    Gravity As Boolean
    ' Object
    ChangedCells As Dictionary
    ' Next
    IncomingShapes As Collection
    Nexts() As Tetris
    ' Hold
    HeldShape As TetrominoShape
    Held As Tetris
    IsHeld As Boolean
    ' Piece
    OverrideShape As TetrominoShape
    DroppingShape As TetrominoShape
    DroppingPiece As Tetromino
    GhostPiece As Tetromino
    ' Level
    Level As Long
    TotalLinesCleared As Long
    LinesCleared As Long
    LinesToLevelUp As Long
    Score As Long
    ' Time
    PieceStayTime As Long
    LastPressedKey As Integer
    LastMoveDownTime As Long
    ' Kick
    KickTable As Variant
    ' States
    Spawn As Boolean
    Ground As Boolean
    Clear As Boolean
    DrawBoard As Boolean
    DrawHold As Boolean
    DrawNexts As Boolean
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
    Let StartButtonText = "S T A R T"
    Call This.Worksheet.Unprotect
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
    Let .Worksheet.Range("F25").Resize(, 3).Value2 = Array(.Level \ 100, _
                                                           .Level \ 10 Mod 10, _
                                                           .Level Mod 10)
End With
End Property
'-----------------------------------------------------------------------------80
Private Property Let TotalLinesCleared(ByVal Lines As Long)
With This
    Let .TotalLinesCleared = Lines
    Let .Worksheet.Range("D22").Resize(, 5).Value2 = Array(.TotalLinesCleared \ 10000 Mod 10, _
                                                           .TotalLinesCleared \ 1000 Mod 10, _
                                                           .TotalLinesCleared \ 100 Mod 10, _
                                                           .TotalLinesCleared \ 10 Mod 10, _
                                                           .TotalLinesCleared Mod 10)
End With
End Property
'-----------------------------------------------------------------------------80
Private Property Let LinesCleared(ByVal Lines As Long)
With This
    Let .LinesCleared = Lines
    Let .Worksheet.Range("F18").Resize(, 3).Value2 = Array(.LinesCleared \ 100, _
                                                           .LinesCleared \ 10 Mod 10, _
                                                           .LinesCleared Mod 10)
    If .LinesCleared >= .LinesToLevelUp Then Let Level = 1 + .Level
End With
End Property
'-----------------------------------------------------------------------------80
Private Property Let LinesToLevelUp(ByVal Lines As Long)
With This
    Let .LinesToLevelUp = Lines
    Let .Worksheet.Range("F20").Resize(, 3).Value2 = Array(.LinesToLevelUp \ 100, _
                                                           .LinesToLevelUp \ 10 Mod 10, _
                                                           .LinesToLevelUp Mod 10)
End With
End Property
'-----------------------------------------------------------------------------80
Private Property Let Score(ByVal Score As Long)
With This
    Let .Score = Score
    Let .Worksheet.Range("B27").Resize(, 7).Value2 = Array(.Score \ 1000000, _
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
    Let This.Worksheet.Range("B29").Value2 = str
End Property
'-----------------------------------------------------------------------------80
Private Property Let StartButtonText(ByVal text As String)
    Let This.Worksheet.Range("B2").Value2 = text
End Property
'-----------------------------------------------------------------------------80
Private Property Let DroppingShape(ByVal DroppingShape As TetrominoShape)
With This
    If DroppingShape = 0 Then .DroppingPiece = Nothing
    Let .DroppingShape = DroppingShape
    Let .Spawn = True
End With
End Property
'-----------------------------------------------------------------------------80
Private Property Set DroppingPiece(ByRef DroppingPiece As Tetromino)
With This
    Dim B As Variant

    If Not .DroppingPiece Is Nothing Then
        For Each B In .DroppingPiece.Blocks
            Set .Matrix(B.x, B.y).Block = Nothing
            Call AddChangedCell(.Matrix(B.x, B.y))
        Next B
    End If

    Set .DroppingPiece = DroppingPiece

    If Not .DroppingPiece Is Nothing Then
        For Each B In .DroppingPiece.Blocks
            Set .Matrix(B.x, B.y).Block = B
            Call AddChangedCell(.Matrix(B.x, B.y))
        Next B
    End If
End With
End Property
'-----------------------------------------------------------------------------80
'Private Property Set DroppingPiece(ByRef Piece As Tetromino)
'With This
'    Dim B As Variant, B2 As Variant
'    Dim InIntersection As Boolean
'    Select Case True
'   '|
'    Case (Not .DroppingPiece Is Nothing) And (Not Piece Is Nothing)
'        For Each B In .DroppingPiece.Blocks
'            Set .Matrix(B.x, B.y).Block = Nothing
'        Next B
'
'        For Each B In Piece.Blocks
'            Set .Matrix(B.x, B.y).Block = B
'        Next B
'        ' For symmetric difference, add them all
'        For Each B In Piece.Blocks
'            Let InIntersection = False
'
'            For Each B2 In .DroppingPiece.Blocks
'                If B.x = B2.x And B.y = B2.y Then ' B in both
'                    ' Intersection
'                    If B.Color <> B2.Color Then
'                        Call AddChangedCell(.Matrix(B.x, B.y))
'                    End If
'                    Let InIntersection = True
'                    Exit For
'                End If
'            Next B2
'
'            If Not InIntersection Then
'                ' In symdiff
'                Call AddChangedCell(.Matrix(B.x, B.y))
'            End If
'        Next B
'
'        For Each B In .DroppingPiece.Blocks
'            Let InIntersection = False
'
'            For Each B2 In Piece.Blocks
'                If B.x = B2.x And B.y = B2.y Then ' B in both
'                    ' Intersection
'                    If B.Color <> B2.Color Then
'                        Call AddChangedCell(.Matrix(B.x, B.y))
'                    End If
'                    Let InIntersection = True
'                    Exit For
'                End If
'            Next B2
'
'            If Not InIntersection Then
'                ' In symdiff
'                Call AddChangedCell(.Matrix(B.x, B.y))
'            End If
'        Next B
'    Case (Not .DroppingPiece Is Nothing) And (Piece Is Nothing)
'        For Each B In .DroppingPiece.Blocks
'            Set .Matrix(B.x, B.y).Block = Nothing
'            Call AddChangedCell(.Matrix(B.x, B.y))
'        Next B
'    Case (.DroppingPiece Is Nothing) And (Not Piece Is Nothing)
'        For Each B In Piece.Blocks
'            Set .Matrix(B.x, B.y).Block = B
'            Call AddChangedCell(.Matrix(B.x, B.y))
'        Next B
'   '|
'    End Select
'
'    Set .DroppingPiece = Piece
'End With
'End Property
'-----------------------------------------------------------------------------80
Private Property Set GhostPiece(ByRef Piece As Tetromino)
With This
    Dim B As Variant, U As Range
    
    If Not .GhostPiece Is Nothing Then
        For Each B In .GhostPiece.Blocks
            If U Is Nothing Then
                Set U = .Matrix(B.x, B.y).Range
            Else
                Set U = Application.Union(.Matrix(B.x, B.y).Range, U)
            End If
            'Call .Matrix(B.x, B.y).Range.BorderAround(, XlBorderWeight.xlMedium, Color:=RGB(31, 31, 31))
'            Let .Matrix(B.x, B.y).Range.Borders(XlBordersIndex.xlEdgeBottom).Color = RGB(31, 31, 31)
'            Let .Matrix(B.x, B.y).Range.Borders(XlBordersIndex.xlEdgeLeft).Color = RGB(31, 31, 31)
'            Let .Matrix(B.x, B.y).Range.Borders(XlBordersIndex.xlEdgeRight).Color = RGB(31, 31, 31)
'            Let .Matrix(B.x, B.y).Range.Borders(XlBordersIndex.xlEdgeTop).Color = RGB(31, 31, 31)
        Next B
        
        Let U.Borders(XlBordersIndex.xlEdgeBottom).Color = RGB(31, 31, 31)
        Let U.Borders(XlBordersIndex.xlEdgeLeft).Color = RGB(31, 31, 31)
        Let U.Borders(XlBordersIndex.xlEdgeRight).Color = RGB(31, 31, 31)
        Let U.Borders(XlBordersIndex.xlEdgeTop).Color = RGB(31, 31, 31)
'        Call U.BorderAround(, XlBorderWeight.xlMedium, Color:=RGB(31, 31, 31))
    End If
    
    Set U = Nothing
    Set .GhostPiece = Piece
    
    If Not .GhostPiece Is Nothing Then
        For Each B In .GhostPiece.Blocks
            If U Is Nothing Then
                Set U = .Matrix(B.x, B.y).Range
            Else
                Set U = Application.Union(.Matrix(B.x, B.y).Range, U)
            End If
            'Call .Matrix(B.x, B.y).Range.BorderAround(, XlBorderWeight.xlMedium, Color:=.DroppingPiece.Color)
            'Let .Matrix(B.x, B.y).Range.Borders(XlBordersIndex.xlEdgeBottom).Color = .DroppingPiece.Color
            'Let .Matrix(B.x, B.y).Range.Borders(XlBordersIndex.xlEdgeLeft).Color = .DroppingPiece.Color
            'Let .Matrix(B.x, B.y).Range.Borders(XlBordersIndex.xlEdgeRight).Color = .DroppingPiece.Color
            'Let .Matrix(B.x, B.y).Range.Borders(XlBordersIndex.xlEdgeTop).Color = .DroppingPiece.Color
        Next B
        
        Let U.Borders(XlBordersIndex.xlEdgeBottom).Color = .DroppingPiece.Color
        Let U.Borders(XlBordersIndex.xlEdgeLeft).Color = .DroppingPiece.Color
        Let U.Borders(XlBordersIndex.xlEdgeRight).Color = .DroppingPiece.Color
        Let U.Borders(XlBordersIndex.xlEdgeTop).Color = .DroppingPiece.Color
        'Call U.BorderAround(, XlBorderWeight.xlMedium, Color:=.DroppingPiece.Color)
    End If
End With
End Property
'=============================================================================80
' Read-Write Property
'=============================================================================80
Public Property Get State() As GameState
    Let State = This.State
End Property
Public Property Let State(ByVal State As GameState)
With This
    Let .State = State
    Select Case .State
        Case GameState.Error:       Let Message = "Error"
        Case GameState.EndManually: Let Message = "Ended"
                                    Let StartButtonText = "S T A R T"
        Case GameState.ToppedOut:   Let Message = "Game Over"
        Case GameState.Paused:      Let Message = "Paused"
                                    Let StartButtonText = "R E S U M E"
        Case GameState.Started:     Let Message = "New Game"
                                    Let StartButtonText = "P A U S S"
        Case GameState.Resumed:     Let Message = "Resumed"
                                    Let StartButtonText = "P A U S S"
    End Select
End With
End Property
'-----------------------------------------------------------------------------80
Public Property Get Gravity() As Boolean
    Let Gravity = This.Gravity
End Property
Private Property Let Gravity(ByVal Toggle As Boolean)
With This
    Let .Gravity = Toggle
    
    If .Gravity Then
        Let Message = "Gravity On"
    Else
        Let Message = "Gravity Off"
    End If
End With
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
    With .Worksheet
        Call .Unprotect
        Call .Range("A1").Select
        Let .Range("A1:AD31").RowHeight = RowHeight
        Let .Range("A1:AD31").ColumnWidth = ColumnWidth
        ' Make the adjacent row and column largest so as to draw less thing on the screen
        Let .Range("AE:AE").ColumnWidth = 255
        Let .Range("32:32").RowHeight = 409
        ' Make the adjacent row and column black
        Let .Range("AE1:AE32").Interior.ColorIndex = 1
        Let .Range("A32:AE32").Interior.ColorIndex = 1
        ' Hide the top two rows according to guideline
        Let .Range("8:9").EntireRow.Hidden = True
        Call .Protect(, True, True, True, , True)
    End With
    
    If Not .IsShowcase Then
        Set .IncomingShapes = New Collection
        Call AddBag
        Set .Held = Tetris.GetTetris(.Worksheet, "D12", 2, 4, True)
        ReDim .Nexts(1 To 6)
        
        Set .Nexts(1) = Tetris.GetTetris(.Worksheet, "X12", 2, 4, True)
        Set .Nexts(2) = Tetris.GetTetris(.Worksheet, "X15", 2, 4, True)
        Set .Nexts(3) = Tetris.GetTetris(.Worksheet, "X18", 2, 4, True)
        Set .Nexts(4) = Tetris.GetTetris(.Worksheet, "X21", 2, 4, True)
        Set .Nexts(5) = Tetris.GetTetris(.Worksheet, "X24", 2, 4, True)
        Set .Nexts(6) = Tetris.GetTetris(.Worksheet, "X27", 2, 4, True)
    End If
    
    ' Draw board, hold, and nexts
    Call DrawAll
    Call .Held.DrawAll
    
    Dim Tet As Variant
    For Each Tet In .Nexts
        Call Tet.DrawAll
    Next Tet
    
    ' Some settings
    
    Let .KickTable = GetKickTable
    Let Level = 1
    Let State = GameState.Started
    Let .Gravity = True
    Let .FirstNonEmptyRow = .Height
    Call GameLoop
End With
End Sub
'-----------------------------------------------------------------------------80
Private Sub GameLoop()
    'On Error GoTo Err
    
    Dim Time As Long, LastRepeatTime As Long, LastKeyPressedTime As Long
    Dim LastMoveDownTime As Long: Let LastMoveDownTime = 1000 * Timer()
    
    With This
   '| Decrease indentation level
    Do While .State >= GameState.Paused ' If active
   '| Decrease indentation level
    DoEvents
    
    If .State = GameState.Paused Then ' If paused
        Select Case True
            Case &H8000 And GetAsyncKeyState(vbKeyR):      Let State = GameState.Resumed
            Case &H8000 And GetAsyncKeyState(vbKeyEscape): Let State = GameState.EndManually
        End Select
    Else ' .GameState > 0
        Let Time = 1000 * Timer()
        
        If .DrawBoard Then
            Call DrawChanged
            Let .DrawBoard = False
        End If
        
        If .DrawNexts Then
            Dim Tet As Variant
            For Each Tet In .Nexts
                Call Tet.DrawChanged
            Next Tet
            
            Let .DrawNexts = False
        End If
            
        If .DrawHold Then
            Call .Held.DrawChanged
            Let .DrawHold = False
        End If
        
        If .Clear Then
            Let .Clear = False
            Call ClearLines
        ElseIf .Ground Then
            Let .Ground = False
            Call Ground
            Let .Clear = True
        ElseIf .OverrideShape <> 0 Then
            Call Spawn(.OverrideShape)
            Let .OverrideShape = 0
        ElseIf .DroppingShape = 0 Then
            Call SpawnNext
        ElseIf Time - LastMoveDownTime >= .PieceStayTime Then
            If .Gravity Then Call GravityDrop
            Let LastMoveDownTime = Time
        Else
            Select Case True 'KeyCodeConstants
           '| Decrease indentation level
            ' Space to hard drop.
            Case &H8000 And GetAsyncKeyState(vbKeySpace)
                If .LastPressedKey <> vbKeySpace Then
                    Let .LastPressedKey = vbKeySpace
                    Call HardDropKey_Pressed
                End If
            ' Shift and C are to hold.
            Case &H8000 And GetAsyncKeyState(vbKeyShift)
                If .LastPressedKey <> vbKeyShift Then
                    Let .LastPressedKey = vbKeyShift
                    Call HoldKey_Pressed
                End If
            Case &H8000 And GetAsyncKeyState(vbKeyC)
                If .LastPressedKey <> vbKeyC Then
                    Let .LastPressedKey = vbKeyC
                    Call HoldKey_Pressed
                End If
            ' Up arrow and X are to rotate clockwise.
            Case &H8000 And GetAsyncKeyState(vbKeyUp)
                If .LastPressedKey <> vbKeyUp Then
                    Let .LastPressedKey = vbKeyUp
                    Call RotateKey_Pressed(-1)
                End If
            Case &H8000 And GetAsyncKeyState(vbKeyX)
                If .LastPressedKey <> vbKeyX Then
                    Let .LastPressedKey = vbKeyX
                    Call RotateKey_Pressed(-1)
                End If
            ' Ctrl and Z are to rotate counterclockwise.
            Case &H8000 And GetAsyncKeyState(vbKeyControl)
                If .LastPressedKey <> vbKeyControl Then
                    Let .LastPressedKey = vbKeyControl
                    Call RotateKey_Pressed(1)
                End If
            Case &H8000 And GetAsyncKeyState(vbKeyZ)
                If .LastPressedKey <> vbKeyZ Then
                    Let .LastPressedKey = vbKeyZ
                    Call RotateKey_Pressed(1)
                End If
            ' Left, right, and down arrows are the same as on the console.
            Case &H8000 And GetAsyncKeyState(vbKeyLeft)
                If .LastPressedKey <> vbKeyLeft Then
                    Let LastKeyPressedTime = 1000 * Timer()
                    Let .LastPressedKey = vbKeyLeft
                    Call MoveLeftKey_Pressed
                Else
                    Let Time = 1000 * Timer()
                    If Time - LastKeyPressedTime >= RepeatDelay Then
                        If Time - LastRepeatTime >= ReciprocalOfRepeatRate Then
                            Let LastRepeatTime = Time
                            Call MoveLeftKey_Pressed
                        End If
                    End If
                End If
            Case &H8000 And GetAsyncKeyState(vbKeyRight)
                If .LastPressedKey <> vbKeyRight Then
                    Let LastKeyPressedTime = 1000 * Timer()
                    Let .LastPressedKey = vbKeyRight
                    Call MoveRightKey_Pressed
                Else
                    Let Time = 1000 * Timer()
                    If Time - LastKeyPressedTime >= RepeatDelay Then
                        If Time - LastRepeatTime >= ReciprocalOfRepeatRate Then
                            Let LastRepeatTime = Time
                            Call MoveRightKey_Pressed
                        End If
                    End If
                End If
            Case &H8000 And GetAsyncKeyState(vbKeyDown)
                If .LastPressedKey <> vbKeyDown Then
                    Let LastKeyPressedTime = 1000 * Timer()
                    Let .LastPressedKey = vbKeyDown
                    Call SoftDropKey_Pressed
                Else
                    Let Time = 1000 * Timer()
                    If Time - LastKeyPressedTime >= RepeatDelay Then
                        If Time - LastRepeatTime >= ReciprocalOfRepeatRate Then
                            Let LastRepeatTime = Time
                            Call SoftDropKey_Pressed
                        End If
                    End If
                End If
            ' Spawn
            Case &H8000 And GetAsyncKeyState(vbKey1)
                If .LastPressedKey <> vbKey1 Then
                    Let .LastPressedKey = vbKey1
                    Let .OverrideShape = IShape
                End If
            Case &H8000 And GetAsyncKeyState(vbKey2)
                If .LastPressedKey <> vbKey2 Then
                    Let .LastPressedKey = vbKey2
                    Let .OverrideShape = ZShape
                End If
            Case &H8000 And GetAsyncKeyState(vbKey3)
                If .LastPressedKey <> vbKey3 Then
                    Let .LastPressedKey = vbKey3
                    Let .OverrideShape = OShape
                End If
            Case &H8000 And GetAsyncKeyState(vbKey4)
                If .LastPressedKey <> vbKey4 Then
                    Let .LastPressedKey = vbKey4
                    Let .OverrideShape = TShape
                End If
            Case &H8000 And GetAsyncKeyState(vbKey5)
                If .LastPressedKey <> vbKey5 Then
                    Let .LastPressedKey = vbKey5
                    Let .OverrideShape = SShape
                End If
            Case &H8000 And GetAsyncKeyState(vbKey6)
                If .LastPressedKey <> vbKey6 Then
                    Let .LastPressedKey = vbKey6
                    Let .OverrideShape = LShape
                End If
            Case &H8000 And GetAsyncKeyState(vbKey7)
                If .LastPressedKey <> vbKey7 Then
                    Let .LastPressedKey = vbKey7
                    Let .OverrideShape = JShape
                End If
            Case &H8000 And GetAsyncKeyState(vbKeyG)
                If .LastPressedKey <> vbKeyG Then
                    Let .LastPressedKey = vbKeyG
                    Let Gravity = Not Gravity
                End If
            ' Pause
            Case &H8000 And GetAsyncKeyState(vbKeyP)
                Let State = GameState.Paused
            ' End
            Case &H8000 And GetAsyncKeyState(vbKeyEscape)
                Let State = GameState.EndManually
            Case Else
                If .LastPressedKey <> 0 Then Let .LastPressedKey = 0
           '| Decrease indentation level
            End Select
        End If
    End If
   '| Decrease indentation level
    Loop
   '| Decrease indentation level
    End With
    
    Exit Sub
Err:
    Let State = GameState.Error
End Sub
'-----------------------------------------------------------------------------80
Private Sub InitializeCells()
With This
    Dim x As Long, y As Long
    ReDim .Matrix(1 To .Height, 1 To .Width)
    
    For x = 1 To .Height
        For y = 1 To .Width
            Set .Matrix(x, y) = New Cell
            Let .Matrix(x, y).BackgroundColor = BackgroundColor
            Set .Matrix(x, y).Range = .Board(x, y)
        Next y
    Next x
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
        Call .BorderAround(, XlBorderWeight.xlMedium, , RGB(255, 255, 0))
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
Private Function IsPieceValid(ByRef Piece As Tetromino) As Boolean
With This
    Let IsPieceValid = True
    
    Dim B As Variant
    For Each B In Piece.Blocks
        If B.x > .Height Or B.y < 1 Or B.y > .Width Then
            Let IsPieceValid = False
            Exit For
        ElseIf .Matrix(B.x, B.y).State = CellState.GroundedBlock Then
            Let IsPieceValid = False
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
Private Sub SpawnNext()
With This
    If Spawn(.IncomingShapes(1)) Then
        Let .DroppingShape = .IncomingShapes(1) ' Dequeue
        Call .IncomingShapes.Remove(1)
        If .IncomingShapes.Count = 6 Then Call AddBag
        
        Dim i As Long
        For i = 1 To 6
             Call .Nexts(i).Spawn(.IncomingShapes(i))
        Next i
        
        Set GhostPiece = GetGhostPiece()
        Let .DrawNexts = True
    End If
End With
End Sub
'-----------------------------------------------------------------------------80
Public Function Spawn(ByVal Shape As TetrominoShape) As Boolean
With This
    Dim TestPiece As Tetromino
    Set TestPiece = Tetromino.GetTetromino(Shape, IIf(Shape = OShape, 1, 2), IIf(.IsShowcase, 2, 5))
    
    If IsPieceValid(TestPiece) Then
        Set DroppingPiece = TestPiece
        Let .DrawBoard = True
        Let .DroppingShape = Shape
        Let Spawn = True
    Else
        Let State = -1
    End If
End With
End Function
'-----------------------------------------------------------------------------80
Private Sub HoldKey_Pressed()
With This
    If .IsHeld Then
        Let Message = "Already held"
    Else
        Let .IsHeld = True
        Let .DrawHold = True
        Call Hold
        Let Message = "Hold"
    End If
End With
End Sub
'-----------------------------------------------------------------------------80
Private Sub Hold()
With This
    Dim TempShape As TetrominoShape: Let TempShape = .HeldShape
    Let .HeldShape = .DroppingShape
    
    If TempShape <> 0 Then
        Let .OverrideShape = TempShape
    Else
        Let .DroppingShape = 0
    End If
    
    Call .Held.Spawn(.HeldShape)
End With
End Sub
'=============================================================================80
' Method: Translation
'=============================================================================80
Private Function Translate(ByVal x As Long, ByVal y As Long, _
                           ByVal Piece As Tetromino, _
                           ByVal InPlace As Boolean) As Tetromino
    If Not InPlace Then Set Piece = Piece.Copy
    Call Piece.Translate(x, y)
    If IsPieceValid(Piece) Then Set Translate = Piece
End Function
'-----------------------------------------------------------------------------80
Private Function Move(ByVal x As Long, ByVal y As Long) As Boolean
With This
    Dim TestPiece As Tetromino
    Set TestPiece = Translate(x, y, .DroppingPiece, False)

    If Not TestPiece Is Nothing Then
        Set DroppingPiece = TestPiece
        Let .DrawBoard = True
        Let Move = True
    End If
End With
End Function
'-----------------------------------------------------------------------------80
' Move Left/Right/Up
'-----------------------------------------------------------------------------80
Private Sub MoveLeftKey_Pressed()
    Call Move(0, -1)
    Set GhostPiece = GetGhostPiece()
End Sub
'-----------------------------------------------------------------------------80
Private Sub MoveRightKey_Pressed()
    Call Move(0, 1)
    Set GhostPiece = GetGhostPiece()
End Sub
'-----------------------------------------------------------------------------80
Private Sub MoveUpKey_Pressed()
    Call Move(-1, 0)
End Sub
'-----------------------------------------------------------------------------80
' Soft Drop
'-----------------------------------------------------------------------------80
Private Sub GravityDrop()
    If Not Move(1, 0) Then Let This.Ground = True
End Sub
'-----------------------------------------------------------------------------80
Private Sub SoftDropKey_Pressed()
    If Move(1, 0) Then
        Let Score = 1 + This.Score
    Else
        Let This.Ground = True
    End If
End Sub
'-----------------------------------------------------------------------------80
' Hard Drop
'-----------------------------------------------------------------------------80
Private Sub HardDropKey_Pressed()
With This
    With .DroppingPiece.Copy
        Let Score = 2 * MovePieceToBottom(.Self) + This.Score
        Set DroppingPiece = .Self
    End With
    
    Let .Ground = True
    Let .DrawBoard = True
End With
End Sub
'-----------------------------------------------------------------------------80
Private Function MovePieceToBottom(ByRef Piece As Tetromino) As Long ' Return distance moved
    Do While Not Translate(1, 0, Piece, True) Is Nothing
        Let MovePieceToBottom = 1 + MovePieceToBottom
    Loop
    Call Translate(-1, 0, Piece, True)
End Function
'-----------------------------------------------------------------------------80
Private Sub Ground()
With This
    Dim B As Variant
    For Each B In .DroppingPiece.Blocks
        Set B.Parent = Nothing
        Set .Matrix(B.x, B.y).Block = B
        If B.x < .FirstNonEmptyRow Then Let .FirstNonEmptyRow = B.x
    Next B
    
    Let .IsHeld = False
    Let .DroppingShape = 0
    Set .DroppingPiece = Nothing
End With
End Sub
'-----------------------------------------------------------------------------80
Private Function GetGhostPiece() As Tetromino
With This.DroppingPiece.Copy
    Call MovePieceToBottom(.Self)
    Set GetGhostPiece = .Self
End With
End Function
'=============================================================================80
' Method: Rotation
'=============================================================================80
Private Sub RotateKey_Pressed(ByVal Sign As Long)
    Call Rotate(Sign)
    Set GhostPiece = GetGhostPiece()
End Sub
'-----------------------------------------------------------------------------80
Private Function Rotate(ByVal Sign As Long) As Boolean
With This
    Dim TestPiece As Tetromino: Set TestPiece = .DroppingPiece.Copy
    
    Call TestPiece.Rotate(Sign) ' CCW: Sign = 1, CW: Sign = -1
    
    If IsPieceValid(TestPiece) Then
        Let Rotate = True
    Else ' Kick
        Dim Test As Variant
        For Each Test In .KickTable(.DroppingShape, .DroppingPiece.Stance, TestPiece.Stance)
            Call TestPiece.Translate(Test(0), Test(1))
            
            If IsPieceValid(TestPiece) Then
                Let Rotate = True
                Let Message = "Kick"
                Exit For
            End If
        Next Test
    End If
    
    If Rotate Then
        Set DroppingPiece = TestPiece
        Let .DrawBoard = True
    End If
End With
End Function
'-----------------------------------------------------------------------------80
Private Function GetKickTable() As Variant
    Dim KickTable As Variant: ReDim KickTable(1 To 7, 0 To 3, 0 To 3)
    
    Dim i As Long, j As Long, k As Long
    For i = 1 To 7
        For j = 0 To 3
            For k = 0 To 3
                Let KickTable(i, j, k) = GetKickTests(i, j, k)
            Next k
        Next j
    Next i
    
    Let GetKickTable = KickTable
End Function
'=============================================================================80
' Method: Clearing
'=============================================================================80
Private Sub ClearLines()
With This
    Dim LineCount As Long: Let LineCount = ClearLines_Internal()
    
    If LineCount Then ' <> 0
        Let .DrawBoard = True
        Let .FirstNonEmptyRow = LineCount + .FirstNonEmptyRow
        Let LinesCleared = LineCount + .LinesCleared
        Let TotalLinesCleared = LineCount + .TotalLinesCleared
        Let Score = 100 * Choose(LineCount, 1, 3, 5, 8) * .Level + .Score
        Let Message = Choose(LineCount, "Single", "Double", "Triple", "Tetris!")
    End If
End With
End Sub
'-----------------------------------------------------------------------------80
Private Function ClearLines_Internal() As Long
With This
    Dim x As Long, y As Long, i As Long
    Dim LineCount As Long, RowOfFirstLine As Long
    Dim MoveRow() As Long: ReDim MoveRow(.FirstNonEmptyRow To .Height)
    ' Determine the mapping MoveRow: SourceRow |-> TargetRow or 0,
    ' where SourceRow has at least one block and is not a line
    For x = .Height To .FirstNonEmptyRow Step -1
        For y = 1 To .Width
            If .Matrix(x, y).State <> CellState.GroundedBlock Then Exit For
        Next y
        
        If y = 1 + .Width Then
            For i = .FirstNonEmptyRow To x - 1
                Let MoveRow(i) = 1 + MoveRow(i)
            Next i
            
            If RowOfFirstLine = 0 Then Let RowOfFirstLine = x
            Let LineCount = 1 + LineCount
            Let MoveRow(x) = 0
        Else
            Let MoveRow(x) = x + MoveRow(x)
        End If
    Next x
    
    If LineCount <> 0 Then
        ' Move rows down
          For x = RowOfFirstLine - 1 To .FirstNonEmptyRow Step -1
            If MoveRow(x) <> 0 Then
                For y = 1 To .Width
                    With .Matrix(MoveRow(x), y)
                        If .Color <> This.Matrix(x, y).Color Then
                            Set .Block = This.Matrix(x, y).Block
                            Call AddChangedCell(.Self)
                        End If
                    End With
                Next y
            End If
        Next x
        ' Clear rows
        For x = .FirstNonEmptyRow To -1 + LineCount + .FirstNonEmptyRow
            For y = 1 To .Width
                With .Matrix(x, y)
                    If .State = CellState.GroundedBlock Then
                        Set .Block = Nothing
                        Call AddChangedCell(.Self)
                    End If
                End With
            Next y
        Next x
    End If
    
    Let ClearLines_Internal = LineCount
End With
End Function
'=============================================================================80
' Helper Function
'=============================================================================80
' http://www.cpearson.com/excel/ShuffleArray.aspx
Private Function ShuffleArray(InArray As Variant) As Variant
    Dim N As Long, k As Long, Temp As Long
    Dim Arr() As Variant: ReDim Arr(LBound(InArray) To UBound(InArray))
    
    Call Randomize
    
    For N = LBound(InArray) To UBound(InArray)
        Let Arr(N) = InArray(N)
    Next N
    
    For N = LBound(InArray) To UBound(InArray)
        Let k = CLng(((UBound(InArray) - N) * Rnd) + N)
        Let Temp = Arr(N)
        Let Arr(N) = Arr(k)
        Let Arr(k) = Temp
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
    'Let .DisplayFormulaBar = Not Toggle
    'Let .DisplayScrollBars = Not Toggle
    'Let .DisplayStatusBar = Not Toggle
    Let .EnableAnimations = Not Toggle
    Let .EnableEvents = Not Toggle
    Let .PrintCommunication = Not Toggle
    'Let .ScreenUpdating = Not toggle
    'Call .ExecuteExcel4Macro("Show.Toolbar(""Ribbon""," & CStr(Not Toggle) & ")")
    
    With .ActiveWindow
        Let .DisplayGridlines = Not Toggle
        'Let .DisplayHeadings = Not Toggle
        'Let .DisplayWorkbookTabs = Not Toggle
    End With
End With
End Sub
