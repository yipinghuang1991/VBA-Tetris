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
'=============================================================================80
' Variable Container
'=============================================================================80
Private Type Variable
' Set Up
    GameState As Long ' -1: Paused, 0: Not started, 1: Ongoing
    Worksheet As Worksheet
    IsShowcase As Boolean
    Height As Long
    Width As Long
    Cell_1_1_Address As String
    Board As Range
    Matrix() As Cell
    Rows() As Variant
    
    ChangedCells As Dictionary
    BoundaryBlocks As Collection
    
    FirstNonEmptyRow As Long
    
    ' Next
    IncomingShapes As Collection
    Nexts() As Tetris
        
    ' Hold
    HeldShape As TetrominoShape
    Held As Tetris
    IsHeld As Boolean
    
    FallingShape As TetrominoShape
    FallingPiece As Tetromino
    TestPiece As Tetromino
    
    Level As Long
    Score As Long
    PieceStayTime As Double
    LastPressedKey As Integer
    LastKeyPressedTime As Single
    
    LastMoveDownTime As Single
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
    Let This.Height = Height + IIf(This.IsShowcase, 0, 3)
End Property
'-----------------------------------------------------------------------------80
Public Property Let Width(ByVal Width As Long)
    Let This.Width = Width + IIf(This.IsShowcase, 0, 2)
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
    
    Let .Worksheet.Range("E15").Resize(, 3).Value2 = _
                        Array(.Level \ 100, (.Level \ 10) Mod 10, .Level Mod 10)
End With
End Property
'-----------------------------------------------------------------------------80
Private Property Let Message(ByVal str As String)
With This
    Let .Worksheet.Range("B21").Value2 = str
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
    
    Select Case This.GameState
        Case -1: Let Message = "Paused"
        Case 0: Let Message = "Ended"
        Case 1: Let Message = "Resumed"
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
Private Sub AddChangedCell(ByRef ChangedCell As Variant)
    On Error Resume Next
    Call This.ChangedCells.Add(ChangedCell, Constants.vbNullString)
End Sub
'-----------------------------------------------------------------------------80
Public Sub Start()
With This
    With ActiveWindow
        .DisplayHeadings = False
        .DisplayGridlines = False
    End With
    
    With .Worksheet
        Call .Unprotect
        Call .Range("A1").Select
        With .Range("A1:AA31")
            Let .RowHeight = RowHeight
            Let .ColumnWidth = ColumnWidth
        End With
        ' Make the adjacent row and column largest so as to draw less thing on the screen
        Let .Range("AB:AB").ColumnWidth = 255
        Let .Range("32:32").RowHeight = 409
        ' Make the adjacent row and column black
        Let .Range("AB1:AB32").Interior.ColorIndex = 1
        Let .Range("A32:AA32").Interior.ColorIndex = 1
        ' Hide the top two rows according to guideline
        Let .Range("8:9").EntireRow.Hidden = True
        Call .Protect(, True, True, True, , True)
    End With
    
    Let Level = 1
    Let .FirstNonEmptyRow = .Height
    
    If Not .IsShowcase Then
        Call SetBoundaryBlocks
        Set .IncomingShapes = New Collection
        Call AddBag
        Set .Held = Tetris.GetTetris(.Worksheet, "E11", 2, 4, True)
        ReDim .Nexts(1 To 6) As Tetris
        
        Set .Nexts(1) = Tetris.GetTetris(.Worksheet, "W11", 2, 4, True)
        Set .Nexts(2) = Tetris.GetTetris(.Worksheet, "W14", 2, 4, True)
        Set .Nexts(3) = Tetris.GetTetris(.Worksheet, "W17", 2, 4, True)
        Set .Nexts(4) = Tetris.GetTetris(.Worksheet, "W20", 2, 4, True)
        Set .Nexts(5) = Tetris.GetTetris(.Worksheet, "W23", 2, 4, True)
        Set .Nexts(6) = Tetris.GetTetris(.Worksheet, "W26", 2, 4, True)
    End If
    
    Call DrawAll
    Call .Held.DrawAll
    
    Dim Tet As Variant
    For Each Tet In .Nexts
        Call Tet.DrawAll
    Next Tet
    
    Call SpawnIncoming(True, True)
    Let .GameState = 1
    Let Message = "Started"
    Let .LastMoveDownTime = TimerInMilliseconds()
    Call GameLoop
End With
End Sub
'-----------------------------------------------------------------------------80
Private Sub GameLoop()
With This
    Dim Time As Long, LastRepeatTime As Long
    Const RepeatDelay As Long = 200
    Const ReciprocalOfRepeatRate As Long = 35
    
    Do While .GameState <> 0
        DoEvents
        
        If .GameState = 1 Then
            Let Time = TimerInMilliseconds()
            If Time - .LastMoveDownTime >= .PieceStayTime Then
                Call MoveDown(True, True)
                Let .LastMoveDownTime = Time
            Else
                Select Case True
                    Case GetAsyncKeyState(KeyCodeConstants.vbKeyEscape) And &H8000
                        Let GameState = 0
                    Case GetAsyncKeyState(KeyCodeConstants.vbKeyP) And &H8000
                        Let GameState = -1
                    ' Shift and C are to hold.
                    Case GetAsyncKeyState(KeyCodeConstants.vbKeyShift) And &H8000
                        If .LastPressedKey <> KeyCodeConstants.vbKeyShift Then
                            Let .LastPressedKey = KeyCodeConstants.vbKeyShift
                            Call Hold
                        End If
                    Case GetAsyncKeyState(KeyCodeConstants.vbKeyC) And &H8000
                        If .LastPressedKey <> KeyCodeConstants.vbKeyC Then
                            Let .LastPressedKey = KeyCodeConstants.vbKeyC
                            Call Hold
                        End If
                    ' Up arrow and X are to rotate clockwise.
                    Case GetAsyncKeyState(KeyCodeConstants.vbKeyUp) And &H8000
                        If .LastPressedKey <> KeyCodeConstants.vbKeyUp Then
                            Let .LastKeyPressedTime = TimerInMilliseconds()
                            Let .LastPressedKey = KeyCodeConstants.vbKeyUp
                            Call RotateCW(True, True)
                        End If
                    Case GetAsyncKeyState(KeyCodeConstants.vbKeyX) And &H8000
                        If .LastPressedKey <> KeyCodeConstants.vbKeyX Then
                            Let .LastKeyPressedTime = TimerInMilliseconds()
                            Let .LastPressedKey = KeyCodeConstants.vbKeyX
                            Call RotateCW(True, True)
                        End If
                    ' Ctrl and Z are to rotate counterclockwise.
                    Case GetAsyncKeyState(KeyCodeConstants.vbKeyControl) And &H8000
                        If .LastPressedKey <> KeyCodeConstants.vbKeyControl Then
                            Let .LastKeyPressedTime = TimerInMilliseconds()
                            Let .LastPressedKey = KeyCodeConstants.vbKeyControl
                            Call RotateCCW(True, True)
                        End If
                    Case GetAsyncKeyState(KeyCodeConstants.vbKeyZ) And &H8000
                        If .LastPressedKey <> KeyCodeConstants.vbKeyZ Then
                            Let .LastKeyPressedTime = TimerInMilliseconds()
                            Let .LastPressedKey = KeyCodeConstants.vbKeyZ
                            Call RotateCCW(True, True)
                        End If
                    ' Left, right, and down arrows are the same as on the console.
                    Case GetAsyncKeyState(KeyCodeConstants.vbKeyLeft) And &H8000
                        If .LastPressedKey <> KeyCodeConstants.vbKeyLeft Then
                            Let .LastKeyPressedTime = TimerInMilliseconds()
                            Let .LastPressedKey = KeyCodeConstants.vbKeyLeft
                            Call MoveLeft(True, True)
                        Else
                            Let Time = TimerInMilliseconds()
                            If Time - .LastKeyPressedTime >= RepeatDelay Then
                                If Time - LastRepeatTime >= ReciprocalOfRepeatRate Then
                                    Let LastRepeatTime = Time
                                    Call MoveLeft(True, True)
                                End If
                            End If
                        End If
                    Case GetAsyncKeyState(KeyCodeConstants.vbKeyRight) And &H8000
                        If .LastPressedKey <> KeyCodeConstants.vbKeyRight Then
                            Let .LastKeyPressedTime = TimerInMilliseconds()
                            Let .LastPressedKey = KeyCodeConstants.vbKeyRight
                            Call MoveRight(True, True)
                        Else
                            Let Time = TimerInMilliseconds()
                            If Time - .LastKeyPressedTime >= RepeatDelay Then
                                If Time - LastRepeatTime >= ReciprocalOfRepeatRate Then
                                    Let LastRepeatTime = Time
                                    Call MoveRight(True, True)
                                End If
                            End If
                        End If
                    Case GetAsyncKeyState(KeyCodeConstants.vbKeyDown) And &H8000
                        If .LastPressedKey <> KeyCodeConstants.vbKeyDown Then
                            Let .LastKeyPressedTime = TimerInMilliseconds()
                            Let .LastPressedKey = KeyCodeConstants.vbKeyDown
                            Call MoveDown(True, True)
                        Else
                            Let Time = TimerInMilliseconds()
                            If Time - .LastKeyPressedTime >= RepeatDelay Then
                                If Time - LastRepeatTime >= ReciprocalOfRepeatRate Then
                                    Let LastRepeatTime = Time
                                    Call MoveDown(True, True)
                                End If
                            End If
                        End If
                    Case Else
                        If .LastPressedKey <> 0 Then
                            Let .LastPressedKey = 0
                        End If
                End Select
            End If
        ElseIf .GameState = -1 Then
            If GetAsyncKeyState(KeyCodeConstants.vbKeyR) And &H8000 Then
                Let GameState = 1
            End If
        End If
    Loop
    
    Call Terminate
End With
End Sub
'-----------------------------------------------------------------------------80
Private Sub InitializeCells()
With This
    Dim X As Long, Y As Long
    
    ReDim .Matrix(1 To .Height, 1 To .Width) As Cell
    ReDim .Rows(1 To .Height - 1) As Variant
    
    Dim Arr() As Cell
    
    For X = 1 To .Height - 1
        ReDim Arr(2 To .Width - 1) As Cell
        Let .Rows(X) = Arr
    Next X
    
    For X = 1 To .Height
        For Y = 1 To .Width
            Set .Matrix(X, Y) = New Cell
            Let .Matrix(X, Y).BackgroundColor = BackgroundColor
            Set .Matrix(X, Y).Range = .Board(X, Y)
        Next Y
    Next X
    
    For X = 1 To .Height - 1
        For Y = 2 To .Width - 1
            Set .Rows(X)(Y) = .Matrix(X, Y)
        Next Y
    Next X
End With
End Sub
'-----------------------------------------------------------------------------80
Private Sub SetBoundaryBlocks()
With This
    Set .BoundaryBlocks = New Collection
    
    Dim X As Long
    For X = 1 To .Height
        Call .BoundaryBlocks.Add(Block.GetBlock(X, 1, , , RGB(127, 127, 127)))
        Call .BoundaryBlocks.Add(Block.GetBlock(X, .Width, , , RGB(127, 127, 127)))
    Next X
    
    Dim Y As Long
    For Y = 2 To .Width - 1
        Call .BoundaryBlocks.Add(Block.GetBlock(.Height, Y, , , RGB(127, 127, 127)))
    Next Y
    
    Dim B As Block
    For Each B In .BoundaryBlocks
        Set .Matrix(B.X, B.Y).Block = B
    Next B
End With
End Sub
'-----------------------------------------------------------------------------80
Public Sub DrawAll()
With This
    With .Board
        Let .Borders(XlBordersIndex.xlInsideHorizontal).Weight = XlBorderWeight.xlMedium
        Let .Borders(XlBordersIndex.xlInsideVertical).Weight = XlBorderWeight.xlMedium
        Call .BorderAround(, XlBorderWeight.xlMedium)
    End With
    
    Dim C As Variant
    For Each C In .Matrix
        Call C.Draw
    Next C
End With
End Sub
'-----------------------------------------------------------------------------80
Public Sub DrawChanged()
With This
    If .ChangedCells.Count Then
        Dim C As Variant
        For Each C In .ChangedCells.Keys
            Call C.Draw
        Next C
        
        Set .ChangedCells = New Dictionary
    End If
End With
End Sub
'=============================================================================80
' Method: Tetromino
'=============================================================================80
Private Sub AddBag()
    Dim v As Variant
    For Each v In ShuffleArray(Array(1, 2, 3, 4, 5, 6, 7))
        Call This.IncomingShapes.Add(v)
    Next v
End Sub
'=============================================================================80
' Method: Clearing
'=============================================================================80
Private Sub ClearLines()
With This
    If .FirstNonEmptyRow = .Height Then Exit Sub
    
    Dim X As Long, IsLine As Boolean, C As Variant, Xv As Variant
    
    Dim Lines As Collection: Set Lines = New Collection
    Dim RowMoveDownStep() As Long: ReDim RowMoveDownStep(1 To .Height - 1) As Long
    'Dim MoveRow() As Long: ReDim MoveRow(.FirstNonEmptyRow To .Height - 1) As Long
    
'    For X = .FirstNonEmptyRow To .Height - 1
'        Let MoveRow(X) = X
'    Next X
    
    ' Determine row number of lines
    For X = .Height - 1 To .FirstNonEmptyRow Step -1
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
        Let Message = Lines.Count & " line" & IIf(Lines.Count = 1, Constants.vbNullString, "s") & " cleared"
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
' Method: Gravity Drop
'=============================================================================80

'=============================================================================80
' Method: Movements
'=============================================================================80
Private Sub AcceptTestPiece(Optional ByVal Draw As Boolean)
With This
    Call UnlinkPiece
    Set .FallingPiece = .TestPiece
    Call LinkPiece
    If Draw Then Call DrawChanged
End With
End Sub
'-----------------------------------------------------------------------------80
Private Function RotateCW(Optional ByVal AcceptOnSuccess As Boolean, Optional ByVal DrawOnSuccess As Boolean) As Boolean
With This
    Set .TestPiece = .FallingPiece.Copy
    Call .TestPiece.Rotate(-1)
    If Not IsTestPieceOverlapped() Then
        Let RotateCW = True
        If AcceptOnSuccess Then Call AcceptTestPiece(DrawOnSuccess)
    End If
End With
End Function
'-----------------------------------------------------------------------------80
Private Function RotateCCW(Optional ByVal AcceptOnSuccess As Boolean, Optional ByVal DrawOnSuccess As Boolean) As Boolean
With This
    Set .TestPiece = .FallingPiece.Copy
    Call .TestPiece.Rotate(1)
    If Not IsTestPieceOverlapped() Then
        Let RotateCCW = True
        If AcceptOnSuccess Then Call AcceptTestPiece(DrawOnSuccess)
    End If
End With
End Function
'-----------------------------------------------------------------------------80
Private Function MoveLeft(Optional ByVal AcceptOnSuccess As Boolean, Optional ByVal DrawOnSuccess As Boolean) As Boolean
With This
    Set .TestPiece = .FallingPiece.Copy
    Call .TestPiece.Move(0, -1)
    If Not IsTestPieceOverlapped() Then
        Let MoveLeft = True
        If AcceptOnSuccess Then Call AcceptTestPiece(DrawOnSuccess)
    End If
End With
End Function
'-----------------------------------------------------------------------------80
Private Function MoveRight(Optional ByVal AcceptOnSuccess As Boolean, Optional ByVal DrawOnSuccess As Boolean) As Boolean
With This
    Set .TestPiece = .FallingPiece.Copy
    Call .TestPiece.Move(0, 1)
    If Not IsTestPieceOverlapped() Then
        Let MoveRight = True
        If AcceptOnSuccess Then Call AcceptTestPiece(DrawOnSuccess)
    End If
End With
End Function
'-----------------------------------------------------------------------------80
Private Function MoveDown(Optional ByVal AcceptOnSuccess As Boolean, Optional ByVal DrawOnSuccess As Boolean) As Boolean
With This
    Set .TestPiece = .FallingPiece.Copy
    Call .TestPiece.Move(1, 0)
    If Not IsTestPieceOverlapped() Then
        Let MoveDown = True
        If AcceptOnSuccess Then
            Call AcceptTestPiece(DrawOnSuccess)
        End If
    Else
        Call Ground
        Call SpawnIncoming(True, True)
    End If
End With
End Function
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
    'Call StopGravity
    'Stop
    Let Message = Constants.vbNullString
    Call ClearLines
End With
End Sub
'-----------------------------------------------------------------------------80
Private Function SpawnIncoming(Optional ByVal AcceptOnSuccess As Boolean, Optional ByVal DrawOnSuccess As Boolean) As Boolean
With This
    Set .TestPiece = Tetromino.GetTetromino(.IncomingShapes(1), IIf(.IncomingShapes(1) = TetrominoShape.O, 1, 2), IIf(.IsShowcase, 2, 6))
    
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
        Let GameState = 0
    End If
End With
End Function
'-----------------------------------------------------------------------------80
Public Sub Spawn(Optional ByVal Shape As TetrominoShape)
With This
    If .IsShowcase And (Not .FallingPiece Is Nothing) Then Call UnlinkPiece
    If Shape <> 0 Then Let .FallingShape = Shape
    Set .FallingPiece = Tetromino.GetTetromino(.FallingShape, IIf(.FallingShape = TetrominoShape.O, 1, 2), IIf(.IsShowcase, 2, 6))
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
' Helper Function
'=============================================================================80
Private Function IsTestPieceOverlapped() As Boolean
With This
    Dim B As Block
    For Each B In .TestPiece.Blocks
        If .Matrix(B.X, B.Y).State = 0 Then
            Let IsTestPieceOverlapped = True
            Exit For
        End If
    Next B
End With
End Function
'-----------------------------------------------------------------------------80
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
    Let .DisplayAlerts = Not Toggle
    Let .DisplayFormulaBar = Not Toggle
    Let .DisplayStatusBar = Not Toggle
    Let .EnableAnimations = Not Toggle
    Let .EnableEvents = Not Toggle
    Let .PrintCommunication = Not Toggle
    'Let .ScreenUpdating = Not toggle
    Let .Calculation = IIf(Toggle, xlCalculationManual, _
                       IIf(IsMissing(Calc), xlCalculationAutomatic, Calc))
End With
End Sub
'-----------------------------------------------------------------------------80
Private Function TimerInMilliseconds() As Long
    Let TimerInMilliseconds = 1000 * Timer
End Function
