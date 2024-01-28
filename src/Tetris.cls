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
    FallingShape As TetrominoShape
    FallingPiece As Tetromino
    TestPiece As Tetromino
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
Public Property Get GameState() As Long
    Let GameState = This.GameState
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
    Dim LastMoveDownTime As Long: Let LastMoveDownTime = 1000 * Timer()
    ' -2: End Manually, -1: Topped Out, 0: Paused, 1: Ongoing, 2: Resumed
    With This
   '| Decrease indentation level
    Do While .GameState >= 0
        If .GameState > 0 Then
            Let Time = 1000 * Timer()
            If Time - LastMoveDownTime >= .PieceStayTime Then
                If .Gravity Then Call Translate(1, 0, True, True)
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
                        Let LastKeyPressedTime = 1000 * Timer()
                        Let .LastPressedKey = vbKeyLeft
                        Call Translate(0, -1, True, True)
                    Else
                        Let Time = 1000 * Timer()
                        If Time - LastKeyPressedTime >= RepeatDelay Then
                            If Time - LastRepeatTime >= ReciprocalOfRepeatRate Then
                                Let LastRepeatTime = Time
                                Call Translate(0, -1, True, True)
                            End If
                        End If
                    End If
                Case GetAsyncKeyState(vbKeyRight) And &H8000
                    If .LastPressedKey <> vbKeyRight Then
                        Let LastKeyPressedTime = 1000 * Timer()
                        Let .LastPressedKey = vbKeyRight
                        Call Translate(0, 1, True, True)
                    Else
                        Let Time = 1000 * Timer()
                        If Time - LastKeyPressedTime >= RepeatDelay Then
                            If Time - LastRepeatTime >= ReciprocalOfRepeatRate Then
                                Let LastRepeatTime = Time
                                Call Translate(0, 1, True, True)
                            End If
                        End If
                    End If
                Case GetAsyncKeyState(vbKeyDown) And &H8000
                    If .LastPressedKey <> vbKeyDown Then
                        Let LastKeyPressedTime = 1000 * Timer()
                        Let .LastPressedKey = vbKeyDown
                        Call SoftDrop
                    Else
                        Let Time = 1000 * Timer()
                        If Time - LastKeyPressedTime >= RepeatDelay Then
                            If Time - LastRepeatTime >= ReciprocalOfRepeatRate Then
                                Let LastRepeatTime = Time
                                Call SoftDrop
                            End If
                        End If
                    End If
                ' Spawn
                Case GetAsyncKeyState(vbKey1) And &H8000
                    If .LastPressedKey <> vbKey1 Then
                        Let .LastPressedKey = vbKey1
                        Call ForceSpawn(IShape)
                    End If
                Case GetAsyncKeyState(vbKey2) And &H8000
                    If .LastPressedKey <> vbKey2 Then
                        Let .LastPressedKey = vbKey2
                        Call ForceSpawn(ZShape)
                    End If
                Case GetAsyncKeyState(vbKey3) And &H8000
                    If .LastPressedKey <> vbKey3 Then
                        Let .LastPressedKey = vbKey3
                        Call ForceSpawn(OShape)
                    End If
                Case GetAsyncKeyState(vbKey4) And &H8000
                    If .LastPressedKey <> vbKey4 Then
                        Let .LastPressedKey = vbKey4
                        Call ForceSpawn(TShape)
                    End If
                Case GetAsyncKeyState(vbKey5) And &H8000
                    If .LastPressedKey <> vbKey5 Then
                        Let .LastPressedKey = vbKey5
                        Call ForceSpawn(SShape)
                    End If
                Case GetAsyncKeyState(vbKey6) And &H8000
                    If .LastPressedKey <> vbKey6 Then
                        Let .LastPressedKey = vbKey6
                        Call ForceSpawn(LShape)
                    End If
                Case GetAsyncKeyState(vbKey7) And &H8000
                    If .LastPressedKey <> vbKey7 Then
                        Let .LastPressedKey = vbKey7
                        Call ForceSpawn(JShape)
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
                         Optional ByRef PieceToLink As Tetromino, _
                         Optional ByVal Draw As Boolean)
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
    
    If Draw Then Call DrawChanged
End With
End Sub
'-----------------------------------------------------------------------------80
Private Sub LinkPiece(ByRef Piece As Tetromino)
With This
    Dim B As Variant
    For Each B In Piece.Blocks
        Set .Matrix(B.x, B.y).Block = B
        Call AddChangedCell(.Matrix(B.x, B.y))
    Next B
End With
End Sub
'-----------------------------------------------------------------------------80
Private Sub UnlinkPiece(ByRef Piece As Tetromino)
With This
    If Not Piece Is Nothing Then
        Dim B As Variant
        For Each B In Piece.Blocks
            Set .Matrix(B.x, B.y).Block = Nothing
            Call AddChangedCell(.Matrix(B.x, B.y))
        Next B
    End If
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
Private Function SpawnIncoming(Optional ByVal AcceptOnSuccess As Boolean, _
                               Optional ByVal DrawOnSuccess As Boolean) As Boolean
With This
    Set .TestPiece = Tetromino.GetTetromino(.IncomingShapes(1), IIf(.IncomingShapes(1) = OShape, 1, 2), IIf(.IsShowcase, 2, 5))
    
    If IsPieceValid(.TestPiece) Then
        Let SpawnIncoming = True
        If AcceptOnSuccess Then
            Let .FallingShape = .IncomingShapes(1)
            Call .IncomingShapes.Remove(1) ' Queue
            If .IncomingShapes.Count = 6 Then Call AddBag
        
            Dim i As Long
            For i = 1 To 6
                Call .Nexts(i).Spawn(.IncomingShapes(i))
            Next i
            
            Call Spawn
        End If
    Else
        Let GameState = -1
    End If
End With
End Function
'-----------------------------------------------------------------------------80
Public Sub Spawn(Optional ByVal Shape As TetrominoShape)
With This
    If .IsShowcase And (Not .FallingPiece Is Nothing) Then Call ReplacePiece(.FallingPiece)
    If Shape <> 0 Then Let .FallingShape = Shape
    Set .FallingPiece = Tetromino.GetTetromino(.FallingShape, IIf(.FallingShape = OShape, 1, 2), IIf(.IsShowcase, 2, 5))
    Call ReplacePiece(, .FallingPiece, True)
End With
End Sub
'-----------------------------------------------------------------------------80
Public Sub ForceSpawn(ByVal Shape As TetrominoShape)
With This
    Call ReplacePiece(.FallingPiece)
    Let .FallingShape = Shape
    Set .FallingPiece = Tetromino.GetTetromino(.FallingShape, IIf(.FallingShape = OShape, 1, 2), IIf(.IsShowcase, 2, 5))
    Call ReplacePiece(, .FallingPiece, True)
End With
End Sub
'-----------------------------------------------------------------------------80
Private Function Hold() As Boolean
With This
    If .IsHeld Then
        Let Message = "Already held"
    Else
        Let .IsHeld = True
        
        Call ReplacePiece(.FallingPiece)
        
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
Private Function Translate(ByVal x As Long, ByVal y As Long, _
                           Optional ByVal AcceptOnSuccess As Boolean, _
                           Optional ByVal DrawOnSuccess As Boolean) As Boolean
With This
    Set .TestPiece = .FallingPiece.Copy
    Call .TestPiece.Translate(x, y)
    If IsPieceValid(.TestPiece) Then
        Let Translate = True
        If AcceptOnSuccess Then _
            Call ReplacePiece(.FallingPiece, .TestPiece, DrawOnSuccess)
    ElseIf x = 1 Then
        Call Ground
    End If
End With
End Function
'-----------------------------------------------------------------------------80
Private Sub SoftDrop()
    If Translate(1, 0, True, True) Then Let Score = 1 + This.Score
End Sub
'-----------------------------------------------------------------------------80
Private Sub HardDrop()
With This
    Dim x As Long
    Set .TestPiece = .FallingPiece.Copy
    
    Do
        Let x = 1 + x
        Call .TestPiece.Translate(1, 0)
    Loop While IsPieceValid(.TestPiece)
    
    Call .TestPiece.Translate(-1, 0)
    Call ReplacePiece(.FallingPiece, .TestPiece, True)
    Call Ground
    Let Score = 2 * (x - 1) + .Score
End With
End Sub
'-----------------------------------------------------------------------------80
Private Sub UpdateGhostPiece()
With This
    Dim x As Long
    
    Call ReplacePiece(.GhostPiece)
    
    Set .GhostPiece = .FallingPiece.Copy

    Do
        Let x = 1 + x
        Call .GhostPiece.Translate(1, 0)
    Loop While IsPieceValid(.GhostPiece)

    Call .GhostPiece.Translate(-1, 0)
    
    Call ReplacePiece(, .GhostPiece)
End With
End Sub
'-----------------------------------------------------------------------------80
Private Sub Ground()
With This
    Let .IsHeld = False
    
    Dim B As Variant
    For Each B In .FallingPiece.Blocks
        Set B.Parent = Nothing
        Set .Matrix(B.x, B.y).Block = B
        If B.x < .FirstNonEmptyRow Then Let .FirstNonEmptyRow = B.x
    Next B
    
    Let Message = vbNullString
    Call ClearLines
    Call SpawnIncoming(True, True)
End With
End Sub
'=============================================================================80
' Method: Rotation
'=============================================================================80
Private Function Rotate(ByVal Sign As Long, _
                        Optional ByVal AcceptOnSuccess As Boolean, _
                        Optional ByVal DrawOnSuccess As Boolean) As Boolean
With This
    Set .TestPiece = .FallingPiece.Copy
    Call .TestPiece.Rotate(Sign) ' CCW: Sign = 1, CW: Sign = -1
    If IsPieceValid(.TestPiece) Then
        Let Rotate = True
        If AcceptOnSuccess Then _
            Call ReplacePiece(.FallingPiece, .TestPiece, DrawOnSuccess)
'    ElseIf Kick() Then
'        Let Rotate = True
'        Let Message = "Wall kick"
'        If AcceptOnSuccess Then Call AcceptTestPiece(DrawOnSuccess)
    Else ' Kick
        Dim Test As Variant
        For Each Test In .KickTable(.FallingShape, .FallingPiece.Stance, .TestPiece.Stance)
            Call .TestPiece.Translate(Test(0), Test(1))
            
            If IsPieceValid(.TestPiece) Then
                Let Rotate = True
                Let Message = "Wall kick"
                If AcceptOnSuccess Then _
                    Call ReplacePiece(.FallingPiece, .TestPiece, DrawOnSuccess)
                Exit For
            End If
        Next Test
    End If
End With
End Function
'-----------------------------------------------------------------------------80
'Private Function Kick() As Boolean
'With This
'    Dim Test As Variant
'    For Each Test In .KickTable(.FallingShape, .FallingPiece.Stance, .TestPiece.Stance)
'        Call .TestPiece.Translate(Test(0), Test(1))
'
'        If IsPieceValid(.TestPiece) Then
'            Let Kick = True
'            Exit For
'        End If
'    Next Test
'End With
'End Function
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
'-----------------------------------------------------------------------------80
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
Private Function GetKickTests(ByVal Shape As TetrominoShape, _
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
'=============================================================================80
' Method: Clearing
'=============================================================================80
Private Sub ClearLines()
With This
    Dim x As Long, y As Long, i As Long
    Dim Lines As Collection: Set Lines = New Collection
    Dim MoveRow() As Long: ReDim MoveRow(.FirstNonEmptyRow To .Height)
    ' Determine the mapping MoveRow: SourceRow |-> TargetRow, where SourceRow has at least one block
    For x = .Height To .FirstNonEmptyRow Step -1
        For y = 1 To .Width
            If .Matrix(x, y).State <> 0 Then Exit For
        Next y
        
        If y = 1 + .Width Then
            For i = .FirstNonEmptyRow To x - 1
                Let MoveRow(i) = 1 + MoveRow(i)
            Next i
            
            Call Lines.Add(x)
            Let MoveRow(x) = 0
        End If
        
        Let MoveRow(x) = x + MoveRow(x)
    Next x
    
    If Lines.Count Then
        ' Move rows down
        For x = Lines(1) - 1 To .FirstNonEmptyRow Step -1
            'If MoveRow(X) <> 0 Then Call MoveRowToRow(X, MoveRow(X))
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
        For x = .FirstNonEmptyRow To -1 + Lines.Count + .FirstNonEmptyRow
            'Call DeleteRow(X)
            For y = 1 To .Width
                With .Matrix(x, y)
                    If .State = 0 Then
                        Set .Block = Nothing
                        Call AddChangedCell(.Self)
                    End If
                End With
            Next y
        Next x
        
        Let .FirstNonEmptyRow = Lines.Count + .FirstNonEmptyRow
        Let LinesCleared = Lines.Count + .LinesCleared
        Let Score = 100 * Choose(Lines.Count, 1, 3, 5, 8) * .Level + .Score
        Let Message = Choose(Lines.Count, "Single", "Double", "Triple", "Tetris!")
    End If
End With
End Sub
'-----------------------------------------------------------------------------80
'Private Sub MoveRowToRow(ByVal X As Long, ByVal ToX As Long)
'With This
'    Dim Y As Long
'    For Y = 1 To .Width
'        With This.Matrix(X, Y).Matrix(ToX, Y)
'            If .Color <> C.Color Then
'                Set .Block = This.Matrix(X, Y).Block
'                'If .State = 0 Then Let .Block.X = ToX
'                Call AddChangedCell(.Self)
'            End If
'        End With
'    Next Y
'End With
'End Sub
'-----------------------------------------------------------------------------80
'Private Sub DeleteRow(ByVal X As Long)
'With This
'    Dim Y As Long
'    For Y = 1 To .Width
'        With .Matrix(X, Y)
'            If .State = 0 Then
'                Set .Block = Nothing
'                Call AddChangedCell(.Self)
'            End If
'        End With
'    Next Y
'End With
'End Sub
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
