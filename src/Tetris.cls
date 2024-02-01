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
' Variable Container
'=============================================================================80
Private Type Variable
' Set Up
    ' -3: Error, -2: End Manually, -1: Topped Out, 0: Paused, 1: Ongoing, 2: Resumed
    GameState As Long
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
    FallingShape As TetrominoShape
    FallingPiece As Tetromino
    GhostPiece As Tetromino
    ' Level
    Level As Long
    LinesCleared As Long
    LinesToLevelUp As Long
    Score As Long
    ' Time
    PieceStayTime As Long
    LastPressedKey As Integer
    LastMoveDownTime As Long
    ' Kick
    KickTable As Variant
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
    Let This.Worksheet.Range("U24").Value2 = str
End Property
'-----------------------------------------------------------------------------80
Public Property Let Gravity(ByVal bool As Boolean)
    Let This.Gravity = bool
    
'    If This.Gravity Then
'        Let Message = "Gravity On"
'    Else
'        Let Message = "Gravity Off"
'    End If
End Property
'=============================================================================80
' Read-Write Property
'=============================================================================80
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
        Let .Range("8:9").EntireRow.Hidden = True
        Call .Protect(, True, True, True, , True)
    End With
    
    If Not .IsShowcase Then
        Set .IncomingShapes = New Collection
        Call AddBag
        Set .Held = Tetris.GetTetris(.Worksheet, "U11", 2, 4, True)
        ReDim .Nexts(1 To 6)
        
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
    Let .KickTable = GetKickTable
    Let Level = 1
    Let GameState = 1
    Let .Gravity = True
    Let .FirstNonEmptyRow = .Height
    Call GameLoop
End With
End Sub
'-----------------------------------------------------------------------------80
Private Sub GameLoop()
    On Error GoTo Err
    
    Dim Time As Long, LastRepeatTime As Long, LastKeyPressedTime As Long
    Dim LastMoveDownTime As Long: Let LastMoveDownTime = 1000 * Timer()
    ' -2: End Manually, -1: Topped Out, 0: Paused, 1: Ongoing, 2: Resumed
    With This
   '| Decrease indentation level
    Do While .GameState >= 0 ' If active
   '| Decrease indentation level
    DoEvents
    
    If .GameState = 0 Then ' If paused
        If GetAsyncKeyState(vbKeyR) And &H8000 Then Let GameState = 2
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
        ElseIf .FallingShape = 0 Then
            Call SpawnNext
        ElseIf Time - LastMoveDownTime >= .PieceStayTime Then
            If .Gravity Then Call GravityDrop
            Let LastMoveDownTime = Time
        Else
            Select Case True 'KeyCodeConstants
           '| Decrease indentation level
            ' Space to hard drop.
            Case GetAsyncKeyState(vbKeySpace) And &H8000
                If .LastPressedKey <> vbKeySpace Then
                    Let .LastPressedKey = vbKeySpace
                    Call HardDropKey_Pressed
                End If
            ' Shift and C are to hold.
            Case GetAsyncKeyState(vbKeyShift) And &H8000
                If .LastPressedKey <> vbKeyShift Then
                    Let .LastPressedKey = vbKeyShift
                    Call HoldKey_Pressed
                End If
            Case GetAsyncKeyState(vbKeyC) And &H8000
                If .LastPressedKey <> vbKeyC Then
                    Let .LastPressedKey = vbKeyC
                    Call HoldKey_Pressed
                End If
            ' Up arrow and X are to rotate clockwise.
            Case GetAsyncKeyState(vbKeyUp) And &H8000
                If .LastPressedKey <> vbKeyUp Then
                    Let .LastPressedKey = vbKeyUp
                    Call RotateKey_Pressed(-1)
                End If
            Case GetAsyncKeyState(vbKeyX) And &H8000
                If .LastPressedKey <> vbKeyX Then
                    Let .LastPressedKey = vbKeyX
                    Call RotateKey_Pressed(-1)
                End If
            ' Ctrl and Z are to rotate counterclockwise.
            Case GetAsyncKeyState(vbKeyControl) And &H8000
                If .LastPressedKey <> vbKeyControl Then
                    Let .LastPressedKey = vbKeyControl
                    Call RotateKey_Pressed(1)
                End If
            Case GetAsyncKeyState(vbKeyZ) And &H8000
                If .LastPressedKey <> vbKeyZ Then
                    Let .LastPressedKey = vbKeyZ
                    Call RotateKey_Pressed(1)
                End If
            ' Left, right, and down arrows are the same as on the console.
            Case GetAsyncKeyState(vbKeyLeft) And &H8000
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
            Case GetAsyncKeyState(vbKeyRight) And &H8000
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
            Case GetAsyncKeyState(vbKeyDown) And &H8000
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
            Case GetAsyncKeyState(vbKey1) And &H8000
                If .LastPressedKey <> vbKey1 Then
                    Let .LastPressedKey = vbKey1
                    Let .OverrideShape = IShape
                End If
            Case GetAsyncKeyState(vbKey2) And &H8000
                If .LastPressedKey <> vbKey2 Then
                    Let .LastPressedKey = vbKey2
                    Let .OverrideShape = ZShape
                End If
            Case GetAsyncKeyState(vbKey3) And &H8000
                If .LastPressedKey <> vbKey3 Then
                    Let .LastPressedKey = vbKey3
                    Let .OverrideShape = OShape
                End If
            Case GetAsyncKeyState(vbKey4) And &H8000
                If .LastPressedKey <> vbKey4 Then
                    Let .LastPressedKey = vbKey4
                    Let .OverrideShape = TShape
                End If
            Case GetAsyncKeyState(vbKey5) And &H8000
                If .LastPressedKey <> vbKey5 Then
                    Let .LastPressedKey = vbKey5
                    Let .OverrideShape = SShape
                End If
            Case GetAsyncKeyState(vbKey6) And &H8000
                If .LastPressedKey <> vbKey6 Then
                    Let .LastPressedKey = vbKey6
                    Let .OverrideShape = LShape
                End If
            Case GetAsyncKeyState(vbKey7) And &H8000
                If .LastPressedKey <> vbKey7 Then
                    Let .LastPressedKey = vbKey7
                    Let .OverrideShape = JShape
                End If
            Case GetAsyncKeyState(vbKeyG) And &H8000
                If .LastPressedKey <> vbKeyG Then
                    Let .LastPressedKey = vbKeyG
                    Let .Gravity = Not .Gravity
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
    End If
   '| Decrease indentation level
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
'Private Sub ReplacePiece(Optional ByRef PieceToUnlink As Tetromino, _
'                         Optional ByRef PieceToLink As Tetromino)
'With This
'    Dim BU As Variant, BL As Variant, IsContainedIn As Boolean
'
'    For Each BU In PieceToUnlink.Blocks
'        Set .Matrix(BU.x, BU.y).Block = Nothing
'        Let IsContainedIn = False
'
'        For Each BL In PieceToLink.Blocks
'            If BU.x = BL.x And BU.y = BL.y Then
'                Let IsContainedIn = True
'                Exit For
'            End If
'        Next BL
'
'        If Not IsContainedIn Then Call AddChangedCell(.Matrix(BU.x, BU.y))
'    Next BU
'
'    For Each BL In PieceToLink.Blocks
'        Set .Matrix(BL.x, BL.y).Block = BL
'
'        For Each BU In PieceToUnlink.Blocks
'            If BU.x = BL.x And BU.y = BL.y Then
'                If BU.Color <> BL.Color Then _
'                    Call AddChangedCell(.Matrix(BL.x, BL.y))
'            End If
'        Next BU
'    Next BL
'
'    Set PieceToUnlink = PieceToLink
'End With
'End Sub
'-----------------------------------------------------------------------------80
Private Sub ReplacePiece(Optional ByRef PieceToUnlink As Tetromino, _
                         Optional ByRef PieceToLink As Tetromino)
With This
    Dim B As Variant
    
    If Not PieceToUnlink Is Nothing Then
        For Each B In PieceToUnlink.Blocks
            Set .Matrix(B.x, B.y).Block = Nothing
            Call AddChangedCell(.Matrix(B.x, B.y))
        Next B
    End If
    
    If Not PieceToLink Is Nothing Then
        For Each B In PieceToLink.Blocks
            Set .Matrix(B.x, B.y).Block = B
            Call AddChangedCell(.Matrix(B.x, B.y))
        Next B
    End If
    
    Set PieceToUnlink = PieceToLink
End With
End Sub
'-----------------------------------------------------------------------------80
Private Function IsPieceValid(ByRef Piece As Tetromino) As Boolean
With This
    Let IsPieceValid = True
    
    Dim B As Variant
    For Each B In Piece.Blocks
        If B.x > .Height Or B.y < 1 Or B.y > .Width Then
            Let IsPieceValid = False
            Exit For
        ElseIf .Matrix(B.x, B.y).State = 0 Then
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
        ' Dequeue
        Let .FallingShape = .IncomingShapes(1)
        Call .IncomingShapes.Remove(1)
        If .IncomingShapes.Count = 6 Then Call AddBag
        
        Dim i As Long
        For i = 1 To 6
             Call .Nexts(i).Spawn(.IncomingShapes(i))
        Next i
        
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
        Call ReplacePiece(.FallingPiece, TestPiece)
        Let .DrawBoard = True
        Let .FallingShape = Shape
        Let Spawn = True
    Else
        Let GameState = -1
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
    Let .HeldShape = .FallingShape
    
    If TempShape <> 0 Then
        Let .OverrideShape = TempShape
    Else
        Let .FallingShape = 0
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
    Set TestPiece = Translate(x, y, .FallingPiece, False)

    If Not TestPiece Is Nothing Then
        Call ReplacePiece(.FallingPiece, TestPiece)
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
End Sub
'-----------------------------------------------------------------------------80
Private Sub MoveRightKey_Pressed()
    Call Move(0, 1)
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
    Let Score = 2 * (HardDropBeforeGrounded(This.FallingPiece)) + This.Score
    Let This.Ground = True
End Sub
'-----------------------------------------------------------------------------80
Private Function HardDropBeforeGrounded(ByRef Piece As Tetromino) As Long
With Piece.Copy
    Dim x As Long
    
    Do While Not Translate(1, 0, .Self, True) Is Nothing
        Let x = 1 + x
    Loop
    
    Call Translate(-1, 0, .Self, True)
    Call ReplacePiece(This.FallingPiece, .Self)
    Let This.DrawBoard = True
    Let HardDropBeforeGrounded = x
End With
End Function
'-----------------------------------------------------------------------------80
Private Sub Ground()
With This
    Dim B As Variant
    For Each B In .FallingPiece.Blocks
        Set B.Parent = Nothing
        Set .Matrix(B.x, B.y).Block = B
        If B.x < .FirstNonEmptyRow Then Let .FirstNonEmptyRow = B.x
    Next B
    
    Let .IsHeld = False
    Let .FallingShape = 0
    Set .FallingPiece = Nothing
End With
End Sub
'=============================================================================80
' Method: Rotation
'=============================================================================80
Private Sub RotateKey_Pressed(ByVal Sign As Long)
    Call Rotate(Sign)
End Sub
'-----------------------------------------------------------------------------80
Private Function Rotate(ByVal Sign As Long) As Boolean
With This
    Dim TestPiece As Tetromino: Set TestPiece = .FallingPiece.Copy
    
    Call TestPiece.Rotate(Sign) ' CCW: Sign = 1, CW: Sign = -1
    
    If IsPieceValid(TestPiece) Then
        Let Rotate = True
    Else ' Kick
        Dim Test As Variant
        For Each Test In .KickTable(.FallingShape, .FallingPiece.Stance, TestPiece.Stance)
            Call TestPiece.Translate(Test(0), Test(1))
            
            If IsPieceValid(TestPiece) Then
                Let Rotate = True
                Let Message = "Kick"
                Exit For
            End If
        Next Test
    End If
    
    If Rotate Then
        Call ReplacePiece(.FallingPiece, TestPiece)
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
            If .Matrix(x, y).State <> 0 Then Exit For
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
                    If .State = 0 Then
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
