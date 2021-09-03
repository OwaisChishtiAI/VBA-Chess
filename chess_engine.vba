'Author: Tom Hedge
'Date: 06/02/2017
'Description: Simulation of a game of Chess
'8 x 8 board with reset button at bottom


Option Explicit


Public Sub CreateWebQuery(Destination As Range, url As String, Optional WebSelectionType As XlWebSelectionType = xlEntirePage, Optional SaveQuery As Boolean, Optional PlainText As Boolean = True)

  '*********************************************************************************'
  '         Builds a web-query object to retrieve information from a web server
  '
  '    Parameters:
  '
  '        Destination
  '          a reference to a cell where the query output will begin
  '
  '        URL
  '          The webpage to get. Should start with "http"
  '
  '        WebSelectionType (xlEntirePage or xlAllTables)
  '          what part of the page should be brought back to Excel.
  '
  '        SaveQuery (True or False)
  '          Indicates if the query object remains in the workbook after running
  '
  '        PlainText (True or False)
  '          Indicates if the quiery results should be plain or include formatting
  '
  '*********************************************************************************'
  
   With Destination.Parent.QueryTables.Add(Connection:="URL;" & url, Destination:=Destination)
        .Name = "WebQuery"
        .RefreshStyle = xlOverwriteCells
        .WebSelectionType = WebSelectionType
        .PreserveFormatting = PlainText
        .BackgroundQuery = False
        .Refresh
        If Not SaveQuery Then .Delete
    End With
    
End Sub



Public Sub webTablesOnCell()
    ' builds a web query looking in the activecell for the URL and returning the tables
    ' from the page in the cell below the active cell
    
    If LCase(Left(ActiveCell.Value, 4)) = "http" Then
      CreateWebQuery ActiveCell.Offset(1), ActiveCell.Value, xlAllTables, False, True
    End If
    
End Sub

Public Sub webPageOnCell()
    ' builds a web query looking in the activecell for the URL and returning the tables
    ' from the page in the cell below the active cell
    
    If LCase(Left(ActiveCell.Value, 4)) = "http" Then
      CreateWebQuery ActiveCell.Offset(1), ActiveCell.Value, xlEntirePage, False, True
    End If
    
End Sub

Sub getPlayersEmails()
    Dim url As String
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQFNtWF9w__Wt3v6E_u7RkpOc6Qy654EndqgdaZBGECLk-18O4mEtIJe7KnD3IT-HV8Ctq_p1TbyWf4/pubhtml?gid=1289366372&single=true"
    
    CreateWebQuery Sheet4.Range("A1"), url
    'Sheet5.Columns(1).Hidden = False
End Sub

Public Function getWhiteEmail() As String
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("FormData")
    'Dim getWhiteEmail As String
    Dim wr As Integer
    wr = sh.Range("C" & Application.Rows.Count).End(xlUp).Row
    getWhiteEmail = Sheet4.Range("C" & wr).Value
End Function

Public Function getBlackEmail() As String
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("FormData")
    'Dim getBlackEmail As String
    Dim br As Integer
    br = sh.Range("D" & Application.Rows.Count).End(xlUp).Row
    getBlackEmail = Sheet4.Range("D" & br).Value
End Function

Sub Send_Email_With_snapshot()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Excel VBA Chess")
    
    Dim wEmail As String
    wEmail = getWhiteEmail
    
    Dim bEmail As String
    bEmail = getBlackEmail
    
    Dim lr As Integer
    'lr = sh.Range("A1:A16").Row
    
    Dim excelfilepath As String
    excelfilepath = Application.ActiveWorkbook.FullName
    
    sh.Range("A1:Q16").Select
    
    With Selection.Parent.MailEnvelope.Item
        .to = wEmail
        .cc = bEmail
        .Subject = "Sample snap"
        .Send
    End With
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=excelfilepath
    Application.DisplayAlerts = True
    
    MsgBox "Email Sent"

End Sub

Private Sub CommandButton1_Click()
    Call getPlayersEmails
    MsgBox "Data Fetching Complete."
End Sub

'Sub to determine what happens each time a cell is selected in the active sheet
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    
    Dim Cell As Range                'Cell being examined
    Dim ClearRange As Range          'Range to be cleared
    Dim Message1 As String           'MessageBox message
    Dim EnPassantMsg As String       'En passant message
    Dim TargetValue As String        'Target cell value
    Dim SelShape As Shape            'Selected shape
    Dim ShpRow As Integer            'Row of selected shape
    Dim ShpCol As Integer            'Column of selected shape
    Dim ShapeNumber As Integer       'Number of shapes in sheet
    Dim DuplicateShape As Shape      'Duplicate shape for Pawn Promotion
    Dim Flag As String               'En Passant, Castling, Pawn Promotion or Check flag
    
    Dim WhitePawn As Pawn            'Define instance of class
    Dim BlackPawn As Pawn            'Define instance of class
    Dim WhiteRook As Rook            'Define instance of class
    Dim BlackRook As Rook            'Define instance of class
    Dim WhiteKnight As Knight        'Define instance of class
    Dim BlackKnight As Knight        'Define instance of class
    Dim WhiteBishop As Bishop        'Define instance of class
    Dim BlackBishop As Bishop        'Define instance of class
    Dim WhiteQueen As Queen          'Define instance of class
    Dim BlackQueen As Queen          'Define instance of class
    Dim WhiteKing As King            'Define instance of class
    Dim BlackKing As King            'Define instance of class
    
    Set WhitePawn = New Pawn            'Instantiate WhitePawn class
    Set BlackPawn = New Pawn            'Instantiate BlackPawn class
    Set WhiteRook = New Rook            'Instantiate WhiteRook class
    Set BlackRook = New Rook            'Instantiate BlackRook class
    Set WhiteKnight = New Knight        'Instantiate WhiteKnight class
    Set BlackKnight = New Knight        'Instantiate BlackKnight class
    Set WhiteBishop = New Bishop        'Instantiate WhiteBishop class
    Set BlackBishop = New Bishop        'Instantiate BlackBishop class
    Set WhiteQueen = New Queen          'Instantiate WhiteQueen class
    Set BlackQueen = New Queen          'Instantiate BlackQueen class
    Set WhiteKing = New King            'Instantiate WhiteKing class
    Set BlackKing = New King            'Instantiate BlackKing class
      
    'To prevent event being triggered during sub
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    'Determine target cell value
    TargetValue = Target.Value
    
    'If more than one cell selected exit
    If Target.Cells.CountLarge > 1 _
        Then
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            Exit Sub
        Else
    End If
    
    'Check to see whether cell selected already in yellow or orange
    For Each Cell In Range("Chessboard")
        If Cell.Range("A1").Interior.ColorIndex = 44 _
            Then
                Set ActivePiece = Cell
                Set ActivePiece2 = Cell
            Else
                If Cell.Range("A1").Interior.ColorIndex = 30 _
                    Then
                        Set CheckedPiece = Cell
                    Else
                        'Do nothing
                End If
        End If
    Next Cell
    
    'Call getPlayersEmails
    If Range("ActivePlayer") = "White" _
        Then
            'Methods if "ActivePlayer = White"
            Player = "White"
            'If a piece is selected
            If Not ActivePiece Is Nothing _
                Then
                    'If target is green ie is a potential movement square
                    If Target.Range("A1").Interior.ColorIndex = 50 Or Target.Interior.ColorIndex = 41 _
                        Then
                            'Cycle through Shapes find one in Target range
                            For Each SelShape In Sheet1.Shapes
                                ShpRow = SelShape.BottomRightCell.Row
                                ShpCol = SelShape.BottomRightCell.Column
                                 'If row and column match Target
                                 If ShpCol = Target.Column And ShpRow = Target.Row _
                                    Then
                                        Set TargetShape = SelShape
                                        
                                        If TargetShape.Name <> "BP1" And TargetShape.Name <> "BP2" And TargetShape.Name <> "BP3" And TargetShape.Name <> "BP4" _
                                            And TargetShape.Name <> "BP5" And TargetShape.Name <> "BP6" And TargetShape.Name <> "BP7" And TargetShape.Name <> "BP8" _
                                            And TargetShape.Name <> "BR1" And TargetShape.Name <> "BR2" And TargetShape.Name <> "BKN1" And TargetShape.Name <> "BKN2" _
                                            And TargetShape.Name <> "BB1" And TargetShape.Name <> "BB2" And TargetShape.Name <> "BQ" _
                                            Then
                                                'Move to Pawn Promotion taken
                                                Call MoveToPPTaken(TargetShape)
                                            Else
                                                'Move to taken pieces
                                                Call MoveToTaken(TargetShape)
                                        End If
                                    Else
                                        'do nothing
                                 End If
                            Next
                            'Cycle through Shapes find one in ActivePiece range
                            For Each SelShape In Sheet1.Shapes
                                 ShpRow = SelShape.BottomRightCell.Row
                                 ShpCol = SelShape.BottomRightCell.Column
                                 'If row and column match
                                 If ShpCol = ActivePiece.Column And ShpRow = ActivePiece.Row _
                                    Then
                                        Set ActiveShape = SelShape
                                        'Move to Target range
                                        Call CenterMe(ActiveShape, Target)
                                        'MsgBox "Done"
                                    Else
                                        'do nothing
                                 End If
                            Next
                            
                            'Change target value to Activepiece value > moving piece
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            
                            Target.Range("A1").Value = ActivePiece.Range("A1").Value
                            'MsgBox "Done"
                            
                            'If active piece is "WP" call ClearMoves Sub
                            If ActivePiece.Range("A1").Value = "WP" _
                                Then
                                    'If checked piece is not nothing
                                    If Not CheckedPiece Is Nothing _
                                        Then
                                            'If checked piece is White King
                                            If CheckedPiece.Value = "WK" _
                                                Then
                                                    'Calc White King possible moves
                                                    Set WhiteKing.Target = CheckedPiece
                                                    WhiteKing.Player = "White"
                                                    WhiteKing.CalcKingMoves
                                                    WhiteKing.EvaluateCheck
                                                    'If White King is in Check
                                                    If WhiteKing.Check = True _
                                                        Then
                                                            'If no moves and in check disallow friendly piece movement
                                                            'Clear Target Move range
                                                            Target.Value = TargetValue
                                                            Call CenterMe(ActiveShape, ActivePiece2)
                                                            If TargetShape Is Nothing _
                                                                Then
                                                                    'Do nothing
                                                                Else
                                                                    Call CenterMe(TargetShape, Target)
                                                            End If
                                                            'Call Clearmoves using ActivePiece and whole of Chessboard for potential moves
                                                            Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                            'Ensure a cell can be reselected
                                                            Range("A1").Select
                                                            'Display message to Player
                                                            Message1 = MsgBox("Move does not break check", vbInformation + vbOKOnly, "Excel VBA Chess")
                                                            
                                                            'Reset ActivePiece and Moves
                                                            Set ActivePiece = Nothing
                                                            Set ActivePiece2 = Nothing
                                                            Set Moves = Nothing
                                                            'Ensure Events are re-enabled
                                                            Application.EnableEvents = True
                                                            Application.ScreenUpdating = True
                                                            Exit Sub
                                                        Else
                                                            'If White King can move or not in Check anymore allow friendly piece movement
                                                            Application.ScreenUpdating = False
                                                            WhitePawn.Player = "White"
                                                            Set WhitePawn.Target = ActivePiece2
                                                            'Turn target to nothing then switch back to "WP"
                                                            Target.Value = ""
                                                            WhitePawn.CalcPawnMoves
                                                            Target.Value = "WP"
                                                            Application.ScreenUpdating = True
                                                            'Call ClearMoves
                                                            Call ClearMoves(ActivePiece2, WhitePawn.PawnMoves)
                                                            'Call EvaluateCheckMoves
                                                            Call ClearMoves(CheckedPiece, Range("chessboard"))
                                                            'Set CheckedPiece to nothing
                                                            Set CheckedPiece = Nothing
                                                            WhitePawn.Player = "White"
                                                            Set WhitePawn.Target = Target
                                                            Call WhitePawn.EvaluateCheckMoves
                                                            If BlackKingCheck = True _
                                                                Then
                                                                    BlackKing.Check = True
                                                                Else
                                                            End If
                                                            'Check if WP promoted
                                                            If Target.Row = 3 _
                                                                Then
                                                                    'If Pawn promoted
                                                                    PawnPromotion.Show
                                                                    'Cycle through Shapes find one in Target range
                                                                    For Each SelShape In Sheet1.Shapes
                                                                         ShpRow = SelShape.BottomRightCell.Row
                                                                         ShpCol = SelShape.BottomRightCell.Column
                                                                         'If row and column match
                                                                         If ShpCol = Target.Column And ShpRow = Target.Row _
                                                                            Then
                                                                                Set ActiveShape = SelShape
                                                                                'Move to taken pieces range
                                                                                
                                                                                'Move to taken pieces
                                                                                Call MoveToTaken(ActiveShape)
                                                                                
                                                                            Else
                                                                                'do nothing
                                                                         End If
                                                                    Next
                                                                    Select Case PromotedPiece
                                                                        Case "Queen"
                                                                            'Initialise Queen value, copy shape and increment
                                                                            Target.Value = "WQ"
                                                                            Set DuplicateShape = Sheet1.Shapes("WQ").Duplicate
                                                                            ShapeNumber = Sheet1.Shapes.Count
                                                                            Sheet1.Shapes(ShapeNumber).Name = "WQ" & WQIncrement
                                                                            Call CenterMe(Sheet1.Shapes(ShapeNumber), Target)
                                                                            Range("A1").Select
                                                                            WQIncrement = WQIncrement + 1
                                                                            'See if Black King in check
                                                                            WhiteQueen.Player = "White"
                                                                            Set WhiteQueen.Target = Target
                                                                            Call WhiteQueen.EvaluateCheckMoves
                                                                            If BlackKingCheck = True _
                                                                                Then
                                                                                    BlackKing.Check = True
                                                                                Else
                                                                            End If
                                                                            Set KingPosition = Nothing
                                                                        Case "Bishop"
                                                                            'Initialise Bishop value, copy shape and increment
                                                                            Target.Value = "WB"
                                                                            Set DuplicateShape = Sheet1.Shapes("WB1").Duplicate
                                                                            ShapeNumber = Sheet1.Shapes.Count
                                                                            Sheet1.Shapes(ShapeNumber).Name = "WB" & WBIncrement
                                                                            Call CenterMe(Sheet1.Shapes(ShapeNumber), Target)
                                                                            Range("A1").Select
                                                                            WBIncrement = WBIncrement + 1
                                                                            'See if Black King in check
                                                                            WhiteBishop.Player = "White"
                                                                            Set WhiteBishop.Target = Target
                                                                            Call WhiteBishop.EvaluateCheckMoves
                                                                            If BlackKingCheck = True _
                                                                                Then
                                                                                    BlackKing.Check = True
                                                                                Else
                                                                            End If
                                                                            Set KingPosition = Nothing
                                                                        Case "Knight"
                                                                            'Initialise Knight value, copy shape and increment
                                                                            Target.Value = "WKN"
                                                                            Set DuplicateShape = Sheet1.Shapes("WKN1").Duplicate
                                                                            ShapeNumber = Sheet1.Shapes.Count
                                                                            Sheet1.Shapes(ShapeNumber).Name = "WKN" & WKNIncrement
                                                                            Call CenterMe(Sheet1.Shapes(ShapeNumber), Target)
                                                                            Range("A1").Select
                                                                            WKNIncrement = WKNIncrement + 1
                                                                            'See if Black King in check
                                                                            WhiteKnight.Player = "White"
                                                                            Set WhiteKnight.Target = Target
                                                                            Call WhiteKnight.EvaluateCheckMoves
                                                                            If BlackKingCheck = True _
                                                                                Then
                                                                                    BlackKing.Check = True
                                                                                Else
                                                                            End If
                                                                            Set KingPosition = Nothing
                                                                        Case "Rook"
                                                                            'Initialise Knight value, copy shape and increment
                                                                            Target.Value = "WR"
                                                                            Set DuplicateShape = Sheet1.Shapes("WR1").Duplicate
                                                                            ShapeNumber = Sheet1.Shapes.Count
                                                                            Sheet1.Shapes(ShapeNumber).Name = "WR" & WRIncrement
                                                                            Call CenterMe(Sheet1.Shapes(ShapeNumber), Target)
                                                                            Range("A1").Select
                                                                            WRIncrement = WRIncrement + 1
                                                                            'See if Black King in check
                                                                            WhiteRook.Player = "White"
                                                                            Set WhiteRook.Target = Target
                                                                            Call WhiteRook.EvaluateCheckMoves
                                                                            If BlackKingCheck = True _
                                                                                Then
                                                                                    BlackKing.Check = True
                                                                                Else
                                                                            End If
                                                                            Set KingPosition = Nothing
                                                                    End Select
                                                                    'MsgBox "Done"
                                                                Else
                                                                    'Do nothing
                                                            End If
                                                            
                                                    End If
                                                Else
                                            End If
                                        Else
                                            'If checked piece is nothing
                                            'Determine whether piece pinned
                                            Call FindKing("White")
                                            Set WhiteKing.Target = KingPosition
                                            
                                            WhiteKing.Player = "White"
                                            ActivePiece2.Value = ""
                                            WhiteKing.CalcKingMoves
                                            WhiteKing.EvaluateCheck
                                            ActivePiece2.Value = "WP"
                                            'If White King is in Check
                                            If WhiteKing.Check = True _
                                                Then
                                                    'If no moves and in check disallow friendly piece movement
                                                    'Clear Target Move range
                                                    Target.Value = TargetValue
                                                    Call CenterMe(ActiveShape, ActivePiece2)
                                                    If TargetShape Is Nothing _
                                                        Then
                                                            'Do nothing
                                                        Else
                                                            Call CenterMe(TargetShape, Target)
                                                    End If
                                                    'Call Clearmoves using ActivePiece and whole of Chessboard for potential moves
                                                    Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                    Call ClearMoves(KingPosition, Range("chessboard"))
                                                    'Ensure a cell can be reselected
                                                    Range("A1").Select
                                                    'Display message to Player
                                                    Message1 = MsgBox("Piece is in an absolute pin and cannot move here", vbInformation + vbOKOnly, "Excel VBA Chess")
                                                    
                                                    'Reset ActivePiece and Moves
                                                    Set ActivePiece = Nothing
                                                    Set ActivePiece2 = Nothing
                                                    Set Moves = Nothing
                                                    Set KingPosition = Nothing
                                                    'Ensure Events are re-enabled
                                                    Application.EnableEvents = True
                                                    Application.ScreenUpdating = True
                                                    Exit Sub
                                                Else
                                                    'Check for En Passant move from left
                                                    If ActivePiece2.Row = 9 And Target.Row = 7 And Target.Offset(0, -1).Value = "BP" _
                                                        Then
                                                            EnPassantMsg = MsgBox("Black Player: Take Pawn 'en passant' from left side of White Player?", vbYesNo + vbQuestion, "Excel VBA Chess")
                                                            If EnPassantMsg = vbYes _
                                                                Then
                                                                    Flag = "EP"
                                                                    For Each SelShape In Sheet1.Shapes
                                                                        ShpRow = SelShape.BottomRightCell.Row
                                                                        ShpCol = SelShape.BottomRightCell.Column
                                                                         'If row and column match Target
                                                                         If ShpCol = Target.Column And ShpRow = Target.Row _
                                                                            Then
                                                                                Set TargetShape = SelShape
                                                                                
                                                                                If TargetShape.Name <> "WP1" And TargetShape.Name <> "WP2" And TargetShape.Name <> "WP3" And TargetShape.Name <> "WP4" _
                                                                                    And TargetShape.Name <> "WP5" And TargetShape.Name <> "WP6" And TargetShape.Name <> "WP7" And TargetShape.Name <> "WP8" _
                                                                                    And TargetShape.Name <> "WR1" And TargetShape.Name <> "WR2" And TargetShape.Name <> "WKN1" And TargetShape.Name <> "WKN2" _
                                                                                    And TargetShape.Name <> "WB1" And TargetShape.Name <> "WB2" And TargetShape.Name <> "WQ" _
                                                                                    Then
                                                                                        'Move to Pawn Promotion taken
                                                                                        Call MoveToPPTaken(TargetShape)
                                                                                    Else
                                                                                        'Move to taken pieces
                                                                                        Call MoveToTaken(TargetShape)
                                                                                End If
                                                                            Else
                                                                                'do nothing
                                                                         End If
                                                                    Next
                                                                    Target.Value = ""
                                                                    'Cycle through Shapes find one in Target range
                                                                    For Each SelShape In Sheet1.Shapes
                                                                        ShpRow = SelShape.BottomRightCell.Row
                                                                        ShpCol = SelShape.BottomRightCell.Column
                                                                         'If row and column match Target
                                                                         If ShpCol = Target.Offset(0, -1).Column And ShpRow = Target.Offset(0, -1).Row _
                                                                            Then
                                                                                Call CenterMe(SelShape, Target.Offset(1, 0))
                                                                            Else
                                                                                'do nothing
                                                                         End If
                                                                         
                                                                    Next
                                                                    
                                                                    'Call Algebraic Turns
                                                                    Call AlgebraicTurns(Target, TargetShape, ActivePiece, ActivePiece2, Flag)
                                                                    
                                                                    'Move pieces
                                                                    Target.Offset(0, -1).Value = ""
                                                                    Target.Offset(1, 0).Value = "BP"
                                                                    ActivePiece2.Value = ""
                                                                    Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                                    
                                                                    'Reset ActivePiece and Moves
                                                                    Set ActivePiece = Nothing
                                                                    Set ActivePiece2 = Nothing
                                                                    Set Moves = Nothing
                                                                    Set KingPosition = Nothing
                                                                    
                                                                    'Increment turns x2
                                                                    Call Increment(Range("Turn").Value)
                                                                    Call Increment(Range("Turn").Value)
                                                                    Range("ActivePlayer").Value = "White"
                                                                    
                                                                    Application.ScreenUpdating = True
                                                                    Application.EnableEvents = True
                                                                    
                                                                    Exit Sub
                                                                Else
                                                                    'Check for en passant move from right
                                                                    If ActivePiece2.Row = 9 And Target.Row = 7 And Target.Offset(0, 1).Value = "BP" _
                                                                        Then
                                                                            EnPassantMsg = MsgBox("Black Player: Take Pawn 'en passant' from right side of White Player?", vbYesNo + vbQuestion, "Excel VBA Chess")
                                                                            If EnPassantMsg = vbYes _
                                                                                Then
                                                                                    Flag = "EP"
                                                                                    For Each SelShape In Sheet1.Shapes
                                                                                        ShpRow = SelShape.BottomRightCell.Row
                                                                                        ShpCol = SelShape.BottomRightCell.Column
                                                                                         'If row and column match Target
                                                                                         If ShpCol = Target.Column And ShpRow = Target.Row _
                                                                                            Then
                                                                                                Set TargetShape = SelShape
                                                                                                
                                                                                                If TargetShape.Name <> "WP1" And TargetShape.Name <> "WP2" And TargetShape.Name <> "WP3" And TargetShape.Name <> "WP4" _
                                                                                                    And TargetShape.Name <> "WP5" And TargetShape.Name <> "WP6" And TargetShape.Name <> "WP7" And TargetShape.Name <> "WP8" _
                                                                                                    And TargetShape.Name <> "WR1" And TargetShape.Name <> "WR2" And TargetShape.Name <> "WKN1" And TargetShape.Name <> "WKN2" _
                                                                                                    And TargetShape.Name <> "WB1" And TargetShape.Name <> "WB2" And TargetShape.Name <> "WQ" _
                                                                                                    Then
                                                                                                        'Move to Pawn Promotion taken
                                                                                                        Call MoveToPPTaken(TargetShape)
                                                                                                    Else
                                                                                                        'Move to taken pieces
                                                                                                        Call MoveToTaken(TargetShape)
                                                                                                End If
                                                                                            Else
                                                                                                'do nothing
                                                                                         End If
                                                                                    Next
                                                                                    Target.Value = ""
                                                                                    'Cycle through Shapes find one in Target range
                                                                                    For Each SelShape In Sheet1.Shapes
                                                                                        ShpRow = SelShape.BottomRightCell.Row
                                                                                        ShpCol = SelShape.BottomRightCell.Column
                                                                                         'If row and column match Target
                                                                                         If ShpCol = Target.Offset(0, 1).Column And ShpRow = Target.Offset(0, 1).Row _
                                                                                            Then
                                                                                                Call CenterMe(SelShape, Target.Offset(1, 0))
                                                                                            Else
                                                                                                'Do nothing
                                                                                         End If
                                                                                         
                                                                                    Next
                                                                                    
                                                                                    'Call Algebraic Turns
                                                                                    Call AlgebraicTurns(Target, TargetShape, ActivePiece, ActivePiece2, Flag)
                                                                                    
                                                                                    'Move pieces
                                                                                    Target.Offset(0, 1).Value = ""
                                                                                    Target.Offset(1, 0).Value = "BP"
                                                                                    ActivePiece2.Value = ""
                                                                                    Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                                                    
                                                                                    'Reset ActivePiece and Moves
                                                                                    Set ActivePiece = Nothing
                                                                                    Set ActivePiece2 = Nothing
                                                                                    Set Moves = Nothing
                                                                                    Set KingPosition = Nothing
                                                                                    
                                                                                    'Increment turns x2
                                                                                    Call Increment(Range("Turn").Value)
                                                                                    Call Increment(Range("Turn").Value)
                                                                                    Range("ActivePlayer").Value = "White"
                                                                                    
                                                                                    Application.ScreenUpdating = True
                                                                                    Application.EnableEvents = True
                                                                                    
                                                                                    Exit Sub
                                                                                Else
                                                                                    'Do nothing
                                                                            End If
                                                                        Else
                                                                    End If
                                                            End If
                                                        Else
                                                            'Check for en passant move from right
                                                            If ActivePiece2.Row = 9 And Target.Row = 7 And Target.Offset(0, 1).Value = "BP" _
                                                                Then
                                                                    EnPassantMsg = MsgBox("Black Player: Take Pawn 'en passant' from right side of White Player?", vbYesNo + vbQuestion, "Excel VBA Chess")
                                                                    If EnPassantMsg = vbYes _
                                                                        Then
                                                                            Flag = "EP"
                                                                            For Each SelShape In Sheet1.Shapes
                                                                                ShpRow = SelShape.BottomRightCell.Row
                                                                                ShpCol = SelShape.BottomRightCell.Column
                                                                                 'If row and column match Target
                                                                                 If ShpCol = Target.Column And ShpRow = Target.Row _
                                                                                    Then
                                                                                        Set TargetShape = SelShape
                                                                                        
                                                                                        If TargetShape.Name <> "WP1" And TargetShape.Name <> "WP2" And TargetShape.Name <> "WP3" And TargetShape.Name <> "WP4" _
                                                                                            And TargetShape.Name <> "WP5" And TargetShape.Name <> "WP6" And TargetShape.Name <> "WP7" And TargetShape.Name <> "WP8" _
                                                                                            And TargetShape.Name <> "WR1" And TargetShape.Name <> "WR2" And TargetShape.Name <> "WKN1" And TargetShape.Name <> "WKN2" _
                                                                                            And TargetShape.Name <> "WB1" And TargetShape.Name <> "WB2" And TargetShape.Name <> "WQ" _
                                                                                            Then
                                                                                                'Move to Pawn Promotion taken
                                                                                                Call MoveToPPTaken(TargetShape)
                                                                                            Else
                                                                                                'Move to taken pieces
                                                                                                Call MoveToTaken(TargetShape)
                                                                                        End If
                                                                                    Else
                                                                                        'do nothing
                                                                                 End If
                                                                            Next
                                                                            Target.Value = ""
                                                                            'Cycle through Shapes find one in Target range
                                                                            For Each SelShape In Sheet1.Shapes
                                                                                ShpRow = SelShape.BottomRightCell.Row
                                                                                ShpCol = SelShape.BottomRightCell.Column
                                                                                 'If row and column match Target
                                                                                 If ShpCol = Target.Offset(0, 1).Column And ShpRow = Target.Offset(0, 1).Row _
                                                                                    Then
                                                                                        Call CenterMe(SelShape, Target.Offset(1, 0))
                                                                                    Else
                                                                                        'Do nothing
                                                                                 End If
                                                                                 
                                                                            Next
                                                                            
                                                                            'Call Algebraic Turns
                                                                            Call AlgebraicTurns(Target, TargetShape, ActivePiece, ActivePiece2, Flag)
                                                                            
                                                                            'Move pieces
                                                                            Target.Offset(0, 1).Value = ""
                                                                            Target.Offset(1, 0).Value = "BP"
                                                                            ActivePiece2.Value = ""
                                                                            Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                                            
                                                                            'Reset ActivePiece and Moves
                                                                            Set ActivePiece = Nothing
                                                                            Set ActivePiece2 = Nothing
                                                                            Set Moves = Nothing
                                                                            Set KingPosition = Nothing
                                                                            
                                                                            'Increment turns x2
                                                                            Call Increment(Range("Turn").Value)
                                                                            Call Increment(Range("Turn").Value)
                                                                            Range("ActivePlayer").Value = "White"
                                                                            
                                                                            Application.ScreenUpdating = True
                                                                            Application.EnableEvents = True
                                                                            
                                                                            Exit Sub
                                                                        Else
                                                                            'Do nothing
                                                                    End If
                                                                Else
                                                            End If
                                                    End If
                                                    'Call ClearMoves
                                                    Set WhitePawn.PawnMoves = Moves
                                                    Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                    'Call EvaluateCheckMoves
                                                    WhitePawn.Player = "White"
                                                    Set WhitePawn.Target = Target
                                                    Call WhitePawn.EvaluateCheckMoves
                                                    If BlackKingCheck = True _
                                                        Then
                                                            BlackKing.Check = True
                                                        Else
                                                    End If
                                                    Set KingPosition = Nothing
                                                    
                                                    'Check if WP promoted
                                                    If Target.Row = 3 _
                                                        Then
                                                            'If Pawn promoted
                                                            PawnPromotion.Show
                                                            'Cycle through Shapes find one in Target range
                                                            For Each SelShape In Sheet1.Shapes
                                                                 ShpRow = SelShape.BottomRightCell.Row
                                                                 ShpCol = SelShape.BottomRightCell.Column
                                                                 'If row and column match
                                                                 If ShpCol = Target.Column And ShpRow = Target.Row _
                                                                    Then
                                                                        Set ActiveShape = SelShape
                                                                        'Move to taken pieces range
                                                                        
                                                                        'Move to taken pieces
                                                                        Call MoveToTaken(ActiveShape)
                                                                    Else
                                                                        'do nothing
                                                                 End If
                                                            Next
                                                            Select Case PromotedPiece
                                                                Case "Queen"
                                                                    'Initialise Queen value, copy shape and increment
                                                                    Target.Value = "WQ"
                                                                    Set DuplicateShape = Sheet1.Shapes("WQ").Duplicate
                                                                    ShapeNumber = Sheet1.Shapes.Count
                                                                    Sheet1.Shapes(ShapeNumber).Name = "WQ" & WQIncrement
                                                                    Call CenterMe(Sheet1.Shapes(ShapeNumber), Target)
                                                                    Range("A1").Select
                                                                    WQIncrement = WQIncrement + 1
                                                                    'See if Black King in check
                                                                    WhiteQueen.Player = "White"
                                                                    Set WhiteQueen.Target = Target
                                                                    Call WhiteQueen.EvaluateCheckMoves
                                                                    If BlackKingCheck = True _
                                                                        Then
                                                                            BlackKing.Check = True
                                                                        Else
                                                                    End If
                                                                    Set KingPosition = Nothing
                                                                Case "Bishop"
                                                                    'Initialise Bishop value, copy shape and increment
                                                                    Target.Value = "WB"
                                                                    Set DuplicateShape = Sheet1.Shapes("WB1").Duplicate
                                                                    ShapeNumber = Sheet1.Shapes.Count
                                                                    Sheet1.Shapes(ShapeNumber).Name = "WB" & WBIncrement
                                                                    Call CenterMe(Sheet1.Shapes(ShapeNumber), Target)
                                                                    Range("A1").Select
                                                                    WBIncrement = WBIncrement + 1
                                                                    'See if Black King in check
                                                                    WhiteBishop.Player = "White"
                                                                    Set WhiteBishop.Target = Target
                                                                    Call WhiteBishop.EvaluateCheckMoves
                                                                    If BlackKingCheck = True _
                                                                        Then
                                                                            BlackKing.Check = True
                                                                        Else
                                                                    End If
                                                                    Set KingPosition = Nothing
                                                                Case "Knight"
                                                                    'Initialise Knight value, copy shape and increment
                                                                    Target.Value = "WKN"
                                                                    Set DuplicateShape = Sheet1.Shapes("WKN1").Duplicate
                                                                    ShapeNumber = Sheet1.Shapes.Count
                                                                    Sheet1.Shapes(ShapeNumber).Name = "WKN" & WKNIncrement
                                                                    Call CenterMe(Sheet1.Shapes(ShapeNumber), Target)
                                                                    Range("A1").Select
                                                                    WKNIncrement = WKNIncrement + 1
                                                                    'See if Black King in check
                                                                    WhiteKnight.Player = "White"
                                                                    Set WhiteKnight.Target = Target
                                                                    Call WhiteKnight.EvaluateCheckMoves
                                                                    If BlackKingCheck = True _
                                                                        Then
                                                                            BlackKing.Check = True
                                                                        Else
                                                                    End If
                                                                    Set KingPosition = Nothing
                                                                Case "Rook"
                                                                    'Initialise Knight value, copy shape and increment
                                                                    Target.Value = "WR"
                                                                    Set DuplicateShape = Sheet1.Shapes("WR1").Duplicate
                                                                    ShapeNumber = Sheet1.Shapes.Count
                                                                    Sheet1.Shapes(ShapeNumber).Name = "WR" & WRIncrement
                                                                    Call CenterMe(Sheet1.Shapes(ShapeNumber), Target)
                                                                    Range("A1").Select
                                                                    WRIncrement = WRIncrement + 1
                                                                    'See if Black King in check
                                                                    WhiteRook.Player = "White"
                                                                    Set WhiteRook.Target = Target
                                                                    Call WhiteRook.EvaluateCheckMoves
                                                                    If BlackKingCheck = True _
                                                                        Then
                                                                            BlackKing.Check = True
                                                                        Else
                                                                    End If
                                                                    Set KingPosition = Nothing
                                                            End Select
                                                        Else
                                                            'Do nothing
                                                    End If
                                            End If
                                    End If
                                    
                                Else
                            
                            End If
                            'MsgBox "Done"
                            'If active piece is "WR" call ClearMoves Sub
                            If ActivePiece.Range("A1").Value = "WR" _
                                Then
                                    'If checked piece is not nothing
                                    If Not CheckedPiece Is Nothing _
                                        Then
                                            'If checked piece is White King
                                            If CheckedPiece.Value = "WK" _
                                                Then
                                                    'Calc White King possible moves
                                                    Set WhiteKing.Target = CheckedPiece
                                                    WhiteKing.Player = "White"
                                                    WhiteKing.CalcKingMoves
                                                    WhiteKing.EvaluateCheck
                                                    'If White King is in Check
                                                    If WhiteKing.Check = True _
                                                        Then
                                                            'If no moves and in check disallow friendly piece movement
                                                            'Clear Target Move range
                                                            Target.Value = TargetValue
                                                            Call CenterMe(ActiveShape, ActivePiece2)
                                                            If TargetShape Is Nothing _
                                                                Then
                                                                    'Do nothing
                                                                Else
                                                                    Call CenterMe(TargetShape, Target)
                                                            End If
                                                            'Call Clearmoves using ActivePiece and whole of Chessboard for potential moves
                                                            Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                            'Ensure a cell can be reselected
                                                            Range("A1").Select
                                                            'Display message to Player
                                                            Message1 = MsgBox("Move does not break check", vbInformation + vbOKOnly, "Excel VBA Chess")
                                                            
                                                            'Reset ActivePiece and Moves
                                                            Set ActivePiece = Nothing
                                                            Set ActivePiece2 = Nothing
                                                            Set Moves = Nothing
                                                            'Ensure Events are re-enabled
                                                            Application.EnableEvents = True
                                                            Application.ScreenUpdating = True
                                                            Exit Sub
                                                        Else
                                                            'If White King can move or not in Check anymore allow friendly piece movement
                                                            Application.ScreenUpdating = False
                                                            WhiteRook.Player = "White"
                                                            Set WhiteRook.Target = ActivePiece2
                                                            'Turn target to nothing then switch back to "WR"
                                                            Target.Value = ""
                                                            WhiteRook.CalcRookMoves
                                                            Target.Value = "WR"
                                                            Application.ScreenUpdating = True
                                                            'Call ClearMoves
                                                            Call ClearMoves(ActivePiece2, WhiteRook.RookMoves)
                                                            'Call EvaluateCheckMoves
                                                            Call ClearMoves(CheckedPiece, Range("chessboard"))
                                                            'Set CheckedPiece to nothing
                                                            Set CheckedPiece = Nothing
                                                            WhiteRook.Player = "White"
                                                            Set WhiteRook.Target = Target
                                                            Call WhiteRook.EvaluateCheckMoves
                                                            If BlackKingCheck = True _
                                                                Then
                                                                    BlackKing.Check = True
                                                                Else
                                                            End If
                                                    End If
                                                Else
                                            End If
                                        Else
                                            'If checked piece is nothing
                                            'Determine whether piece pinned
                                            Call FindKing("White")
                                            Set WhiteKing.Target = KingPosition
                                            
                                            WhiteKing.Player = "White"
                                            ActivePiece2.Value = ""
                                            WhiteKing.CalcKingMoves
                                            WhiteKing.EvaluateCheck
                                            ActivePiece2.Value = "WR"
                                            'If White King is in Check
                                            If WhiteKing.Check = True _
                                                Then
                                                    'If no moves and in check disallow friendly piece movement
                                                    'Clear Target Move range
                                                    Target.Value = TargetValue
                                                    Call CenterMe(ActiveShape, ActivePiece2)
                                                    If TargetShape Is Nothing _
                                                        Then
                                                            'Do nothing
                                                        Else
                                                            Call CenterMe(TargetShape, Target)
                                                    End If
                                                    'Call Clearmoves using ActivePiece and whole of Chessboard for potential moves
                                                    Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                    Call ClearMoves(KingPosition, Range("chessboard"))
                                                    'Ensure a cell can be reselected
                                                    Range("A1").Select
                                                    'Display message to Player
                                                    Message1 = MsgBox("Piece is in an absolute pin and cannot move here", vbInformation + vbOKOnly, "Excel VBA Chess")
                                                    
                                                    'Reset ActivePiece and Moves
                                                    Set ActivePiece = Nothing
                                                    Set ActivePiece2 = Nothing
                                                    Set Moves = Nothing
                                                    Set KingPosition = Nothing
                                                    'Ensure Events are re-enabled
                                                    Application.EnableEvents = True
                                                    Application.ScreenUpdating = True
                                                    Exit Sub
                                                Else
                                                    'Call ClearMoves
                                                    Set WhiteRook.RookMoves = Moves
                                                    Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                    'Call EvaluateCheckMoves
                                                    WhiteRook.Player = "White"
                                                    Set WhiteRook.Target = Target
                                                    Call WhiteRook.EvaluateCheckMoves
                                                    If BlackKingCheck = True _
                                                        Then
                                                            BlackKing.Check = True
                                                        Else
                                                    End If
                                                    Set KingPosition = Nothing
                                            End If
                                    End If
                                Else
                            End If
                            'If active piece is "WKN" call ClearMoves Sub
                            If ActivePiece.Range("A1").Value = "WKN" _
                                Then
                                    'If checked piece is not nothing
                                    If Not CheckedPiece Is Nothing _
                                        Then
                                            'If checked piece is White King
                                            If CheckedPiece.Value = "WK" _
                                                Then
                                                    'Calc White King possible moves
                                                    Set WhiteKing.Target = CheckedPiece
                                                    WhiteKing.Player = "White"
                                                    WhiteKing.CalcKingMoves
                                                    WhiteKing.EvaluateCheck
                                                    'If White King is in Check
                                                    If WhiteKing.Check = True _
                                                        Then
                                                            'If no moves and in check disallow friendly piece movement
                                                            'Clear Target Move range
                                                            Target.Value = TargetValue
                                                            Call CenterMe(ActiveShape, ActivePiece2)
                                                            If TargetShape Is Nothing _
                                                                Then
                                                                    'Do nothing
                                                                Else
                                                                    Call CenterMe(TargetShape, Target)
                                                            End If
                                                            'Call Clearmoves using ActivePiece and whole of Chessboard for potential moves
                                                            Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                            'Ensure a cell can be reselected
                                                            Range("A1").Select
                                                            'Display message to Player
                                                            Message1 = MsgBox("Move does not break check", vbInformation + vbOKOnly, "Excel VBA Chess")
                                                            
                                                            'Reset ActivePiece and Moves
                                                            Set ActivePiece = Nothing
                                                            Set ActivePiece2 = Nothing
                                                            Set Moves = Nothing
                                                            'Ensure Events are re-enabled
                                                            Application.EnableEvents = True
                                                            Application.ScreenUpdating = True
                                                            Exit Sub
                                                        Else
                                                            'If White King can move or not in Check anymore allow friendly piece movement
                                                            Application.ScreenUpdating = False
                                                            WhiteKnight.Player = "White"
                                                            Set WhiteKnight.Target = ActivePiece2
                                                            'Turn target to nothing then switch back to "WKN"
                                                            Target.Value = ""
                                                            WhiteKnight.CalcKnightMoves
                                                            Target.Value = "WKN"
                                                            Application.ScreenUpdating = True
                                                            'Call ClearMoves
                                                            Call ClearMoves(ActivePiece2, WhiteKnight.KnightMoves)
                                                            'Call EvaluateCheckMoves
                                                            Call ClearMoves(CheckedPiece, Range("chessboard"))
                                                            'Set CheckedPiece to nothing
                                                            Set CheckedPiece = Nothing
                                                            WhiteKnight.Player = "White"
                                                            Set WhiteKnight.Target = Target
                                                            Call WhiteKnight.EvaluateCheckMoves
                                                            If BlackKingCheck = True _
                                                                Then
                                                                    BlackKing.Check = True
                                                                Else
                                                            End If
                                                    End If
                                                Else
                                            End If
                                        Else
                                            'If checked piece is nothing
                                            'Determine whether piece pinned
                                            Call FindKing("White")
                                            Set WhiteKing.Target = KingPosition
                                            
                                            WhiteKing.Player = "White"
                                            ActivePiece2.Value = ""
                                            WhiteKing.CalcKingMoves
                                            WhiteKing.EvaluateCheck
                                            ActivePiece2.Value = "WKN"
                                            'If White King is in Check
                                            If WhiteKing.Check = True _
                                                Then
                                                    'If no moves and in check disallow friendly piece movement
                                                    'Clear Target Move range
                                                    Target.Value = TargetValue
                                                    Call CenterMe(ActiveShape, ActivePiece2)
                                                    If TargetShape Is Nothing _
                                                        Then
                                                            'Do nothing
                                                        Else
                                                            Call CenterMe(TargetShape, Target)
                                                    End If
                                                    'Call Clearmoves using ActivePiece and whole of Chessboard for potential moves
                                                    Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                    Call ClearMoves(KingPosition, Range("chessboard"))
                                                    'Ensure a cell can be reselected
                                                    Range("A1").Select
                                                    'Display message to Player
                                                    Message1 = MsgBox("Piece is in an absolute pin and cannot move here", vbInformation + vbOKOnly, "Excel VBA Chess")
                                                    
                                                    'Reset ActivePiece and Moves
                                                    Set ActivePiece = Nothing
                                                    Set ActivePiece2 = Nothing
                                                    Set Moves = Nothing
                                                    Set KingPosition = Nothing
                                                    'Ensure Events are re-enabled
                                                    Application.EnableEvents = True
                                                    Application.ScreenUpdating = True
                                                    Exit Sub
                                                Else
                                                    'Call ClearMoves
                                                    Set WhiteKnight.KnightMoves = Moves
                                                    Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                    'Call EvaluateCheckMoves
                                                    WhiteKnight.Player = "White"
                                                    Set WhiteKnight.Target = Target
                                                    Call WhiteKnight.EvaluateCheckMoves
                                                    If BlackKingCheck = True _
                                                        Then
                                                            BlackKing.Check = True
                                                        Else
                                                    End If
                                                    Set KingPosition = Nothing
                                            End If
                                    End If
                                Else
                            End If
                            'If active piece is "WB" call ClearMoves Sub
                            If ActivePiece.Range("A1").Value = "WB" _
                                Then
                                    'If checked piece is not nothing
                                    If Not CheckedPiece Is Nothing _
                                        Then
                                            'If checked piece is White King
                                            If CheckedPiece.Value = "WK" _
                                                Then
                                                    'Calc White King possible moves
                                                    Set WhiteKing.Target = CheckedPiece
                                                    WhiteKing.Player = "White"
                                                    WhiteKing.CalcKingMoves
                                                    WhiteKing.EvaluateCheck
                                                    'If White King is in Check
                                                    If WhiteKing.Check = True _
                                                        Then
                                                            'If no moves and in check disallow friendly piece movement
                                                            'Clear Target Move range
                                                            Target.Value = TargetValue
                                                            Call CenterMe(ActiveShape, ActivePiece2)
                                                            If TargetShape Is Nothing _
                                                                Then
                                                                    'Do nothing
                                                                Else
                                                                    Call CenterMe(TargetShape, Target)
                                                            End If
                                                            'Call Clearmoves using ActivePiece and whole of Chessboard for potential moves
                                                            Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                            'Ensure a cell can be reselected
                                                            Range("A1").Select
                                                            'Display message to Player
                                                            Message1 = MsgBox("Move does not break check", vbInformation + vbOKOnly, "Excel VBA Chess")
                                                            
                                                            'Reset ActivePiece and Moves
                                                            Set ActivePiece = Nothing
                                                            Set ActivePiece2 = Nothing
                                                            Set Moves = Nothing
                                                            'Ensure Events are re-enabled
                                                            Application.EnableEvents = True
                                                            Application.ScreenUpdating = True
                                                            Exit Sub
                                                        Else
                                                            'If White King can move or not in Check anymore allow friendly piece movement
                                                            Application.ScreenUpdating = False
                                                            WhiteBishop.Player = "White"
                                                            Set WhiteBishop.Target = ActivePiece2
                                                            'Turn target to nothing then switch back to "WB"
                                                            Target.Value = ""
                                                            WhiteBishop.CalcBishopMoves
                                                            Target.Value = "WB"
                                                            Application.ScreenUpdating = True
                                                            'Call ClearMoves
                                                            Call ClearMoves(ActivePiece2, WhiteBishop.BishopMoves)
                                                            'Call EvaluateCheckMoves
                                                            Call ClearMoves(CheckedPiece, Range("chessboard"))
                                                            'Set CheckedPiece to nothing
                                                            Set CheckedPiece = Nothing
                                                            WhiteBishop.Player = "White"
                                                            Set WhiteBishop.Target = Target
                                                            Call WhiteBishop.EvaluateCheckMoves
                                                            If BlackKingCheck = True _
                                                                Then
                                                                    BlackKing.Check = True
                                                                Else
                                                            End If
                                                    End If
                                                Else
                                            End If
                                        Else
                                            'If checked piece is nothing
                                            'Determine whether piece pinned
                                            Call FindKing("White")
                                            Set WhiteKing.Target = KingPosition
                                            
                                            WhiteKing.Player = "White"
                                            ActivePiece2.Value = ""
                                            WhiteKing.CalcKingMoves
                                            WhiteKing.EvaluateCheck
                                            ActivePiece2.Value = "WB"
                                            'If White King is in Check
                                            If WhiteKing.Check = True _
                                                Then
                                                    'If no moves and in check disallow friendly piece movement
                                                    'Clear Target Move range
                                                    Target.Value = TargetValue
                                                    Call CenterMe(ActiveShape, ActivePiece2)
                                                    If TargetShape Is Nothing _
                                                        Then
                                                            'Do nothing
                                                        Else
                                                            Call CenterMe(TargetShape, Target)
                                                    End If
                                                    'Call Clearmoves using ActivePiece and whole of Chessboard for potential moves
                                                    Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                    Call ClearMoves(KingPosition, Range("chessboard"))
                                                    'Ensure a cell can be reselected
                                                    Range("A1").Select
                                                    'Display message to Player
                                                    Message1 = MsgBox("Piece is in an absolute pin and cannot move here", vbInformation + vbOKOnly, "Excel VBA Chess")
                                                    
                                                    'Reset ActivePiece and Moves
                                                    Set ActivePiece = Nothing
                                                    Set ActivePiece2 = Nothing
                                                    Set Moves = Nothing
                                                    Set KingPosition = Nothing
                                                    'Ensure Events are re-enabled
                                                    Application.EnableEvents = True
                                                    Application.ScreenUpdating = True
                                                    Exit Sub
                                                Else
                                                    'Call ClearMoves
                                                    Set WhiteBishop.BishopMoves = Moves
                                                    Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                    'Call EvaluateCheckMoves
                                                    WhiteBishop.Player = "White"
                                                    Set WhiteBishop.Target = Target
                                                    Call WhiteBishop.EvaluateCheckMoves
                                                    If BlackKingCheck = True _
                                                        Then
                                                            BlackKing.Check = True
                                                        Else
                                                    End If
                                                    Set KingPosition = Nothing
                                            End If
                                    End If
                                Else
                            End If
                            'If active piece is "WQ" call ClearMoves Sub
                            If ActivePiece.Range("A1").Value = "WQ" _
                                Then
                                    'If checked piece is not nothing
                                    If Not CheckedPiece Is Nothing _
                                        Then
                                            'If checked piece is White King
                                            If CheckedPiece.Value = "WK" _
                                                Then
                                                    'Calc White King possible moves
                                                    Set WhiteKing.Target = CheckedPiece
                                                    WhiteKing.Player = "White"
                                                    WhiteKing.CalcKingMoves
                                                    WhiteKing.EvaluateCheck
                                                    'If White King is in Check
                                                    If WhiteKing.Check = True _
                                                        Then
                                                            'If no moves and in check disallow friendly piece movement
                                                            'Clear Target Move range
                                                            Target.Value = TargetValue
                                                            Call CenterMe(ActiveShape, ActivePiece2)
                                                            If TargetShape Is Nothing _
                                                                Then
                                                                    'Do nothing
                                                                Else
                                                                    Call CenterMe(TargetShape, Target)
                                                            End If
                                                            'Call Clearmoves using ActivePiece and whole of Chessboard for potential moves
                                                            Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                            'Ensure a cell can be reselected
                                                            Range("A1").Select
                                                            'Display message to Player
                                                            Message1 = MsgBox("Move does not break check", vbInformation + vbOKOnly, "Excel VBA Chess")
                                                            
                                                            'Reset ActivePiece and Moves
                                                            Set ActivePiece = Nothing
                                                            Set ActivePiece2 = Nothing
                                                            Set Moves = Nothing
                                                            'Ensure Events are re-enabled
                                                            Application.EnableEvents = True
                                                            Application.ScreenUpdating = True
                                                            Exit Sub
                                                        Else
                                                            'If White King can move or not in Check anymore allow friendly piece movement
                                                            Application.ScreenUpdating = False
                                                            WhiteQueen.Player = "White"
                                                            Set WhiteQueen.Target = ActivePiece2
                                                            'Turn target to nothing then switch back to "WQ"
                                                            Target.Value = ""
                                                            WhiteQueen.CalcQueenMoves
                                                            Target.Value = "WQ"
                                                            Application.ScreenUpdating = True
                                                            'Call ClearMoves
                                                            Call ClearMoves(ActivePiece2, WhiteQueen.QueenMoves)
                                                            'Call EvaluateCheckMoves
                                                            Call ClearMoves(CheckedPiece, Range("chessboard"))
                                                            'Set CheckedPiece to nothing
                                                            Set CheckedPiece = Nothing
                                                            WhiteQueen.Player = "White"
                                                            Set WhiteQueen.Target = Target
                                                            Call WhiteQueen.EvaluateCheckMoves
                                                            If BlackKingCheck = True _
                                                                Then
                                                                    BlackKing.Check = True
                                                                Else
                                                            End If
                                                    End If
                                                Else
                                            End If
                                        Else
                                            'If checked piece is nothing
                                            'Determine whether piece pinned
                                            Call FindKing("White")
                                            Set WhiteKing.Target = KingPosition
                                            
                                            WhiteKing.Player = "White"
                                            ActivePiece2.Value = ""
                                            WhiteKing.CalcKingMoves
                                            WhiteKing.EvaluateCheck
                                            ActivePiece2.Value = "WQ"
                                            'If White King is in Check
                                            If WhiteKing.Check = True _
                                                Then
                                                    'If no moves and in check disallow friendly piece movement
                                                    'Clear Target Move range
                                                    Target.Value = TargetValue
                                                    Call CenterMe(ActiveShape, ActivePiece2)
                                                    If TargetShape Is Nothing _
                                                        Then
                                                            'Do nothing
                                                        Else
                                                            Call CenterMe(TargetShape, Target)
                                                    End If
                                                    'Call Clearmoves using ActivePiece and whole of Chessboard for potential moves
                                                    Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                    Call ClearMoves(KingPosition, Range("chessboard"))
                                                    'Ensure a cell can be reselected
                                                    Range("A1").Select
                                                    'Display message to Player
                                                    Message1 = MsgBox("Piece is in an absolute pin and cannot move here", vbInformation + vbOKOnly, "Excel VBA Chess")
                                                    
                                                    'Reset ActivePiece and Moves
                                                    Set ActivePiece = Nothing
                                                    Set ActivePiece2 = Nothing
                                                    Set Moves = Nothing
                                                    Set KingPosition = Nothing
                                                    'Ensure Events are re-enabled
                                                    Application.EnableEvents = True
                                                    Application.ScreenUpdating = True
                                                    Exit Sub
                                                Else
                                                    'Call ClearMoves
                                                    Set WhiteQueen.QueenMoves = Moves
                                                    Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                    'Call EvaluateCheckMoves
                                                    WhiteQueen.Player = "White"
                                                    Set WhiteQueen.Target = Target
                                                    Call WhiteQueen.EvaluateCheckMoves
                                                    If BlackKingCheck = True _
                                                        Then
                                                            BlackKing.Check = True
                                                        Else
                                                    End If
                                                    Set KingPosition = Nothing
                                            End If
                                    End If
                                Else
                            End If
                            'If active piece is "WK" call ClearMoves Sub
                            If ActivePiece.Range("A1").Value = "WK" _
                                Then
                                    Set WhiteKing.KingMoves = Moves
                                    Set WhiteKing.ImpossibleMoves = GlobalImpMoves
                                    If WhiteKing.KingMoves Is Nothing And WhiteKing.ImpossibleMoves Is Nothing _
                                        Then
                                            Set ClearRange = Nothing
                                        Else
                                            If Not WhiteKing.KingMoves Is Nothing And WhiteKing.ImpossibleMoves Is Nothing _
                                                Then
                                                    Set ClearRange = WhiteKing.KingMoves
                                                Else
                                                    If WhiteKing.KingMoves Is Nothing And Not WhiteKing.ImpossibleMoves Is Nothing _
                                                        Then
                                                            Set ClearRange = WhiteKing.ImpossibleMoves
                                                        Else
                                                            Set ClearRange = Union(WhiteKing.KingMoves, WhiteKing.ImpossibleMoves)
                                                    End If
                                            End If
                                    End If
                                    'If there are Castlemoves to add to the range clear
                                    Set WhiteKing.CastleMoves = GlobalCastleMoves
                                    If WhiteKing.CastleMoves Is Nothing _
                                        Then
                                            'do nothing
                                        Else
                                            If ClearRange Is Nothing _
                                                Then
                                                    Set ClearRange = WhiteKing.CastleMoves
                                                Else
                                                    Set ClearRange = Union(ClearRange, WhiteKing.CastleMoves)
                                            End If
                                            'Move Rook to protect King
                                            If Target.Column = 8 _
                                                Then
                                                    ActivePiece2.Offset(0, 3).Value = ""
                                                    ActivePiece2.Offset(0, 1).Value = "WR"
                                                    Call CenterMe(Sheet1.Shapes("WR2"), Range("G10"))
                                                    Flag = "Kingside Castling"
                                                Else
                                                    If Target.Column = 4 _
                                                        Then
                                                            ActivePiece2.Offset(0, -4).Value = ""
                                                            ActivePiece2.Offset(0, -1).Value = "WR"
                                                            Call CenterMe(Sheet1.Shapes("WR1"), Range("E10"))
                                                            Flag = "Queenside Castling"
                                                        Else
                                                            'Do nothing
                                                    End If
                                            End If
                                    End If
                                    Call ClearMoves(ActivePiece, ClearRange)
                                    WhiteKingCheck = False
                                    WhiteKingMoved = True
                                    'Set CheckedPiece to nothing
                                    Set CheckedPiece = Nothing
                                Else
                            End If
                            
                            If (WhiteKingCheck = True Or BlackKingCheck = True) Then Flag = "Check"
                            
                            'Populate Algebraic Turns
                            Call AlgebraicTurns(Target, TargetShape, ActivePiece, ActivePiece2, Flag)
                            
                            'Change activepiece2 value to "" > moving piece
                            ActivePiece2.Range("A1").Value = ""
                            Set ActivePiece = Nothing
                            Set ActivePiece2 = Nothing
                            Set Moves = Nothing
                            Call Send_Email_With_snapshot
                            'MsgBox "White"
                            'Change Active Player to "Black"
                            Range("ActivePlayer").Value = "Black"
                            'MsgBox "Done"
                            
                        Else
                            'If target square is yellow ie is selected square > deselect
                            If Target.Range("A1").Interior.ColorIndex = 44 Or Target.Interior.ColorIndex = 30 _
                                Then
                                    'If WP call ClearMovesSub
                                    If ActivePiece.Range("A1").Value = "WP" _
                                        Then
                                            Set WhitePawn.PawnMoves = Moves
                                            Call ClearMoves(ActivePiece, WhitePawn.PawnMoves)
                                        Else
                                    End If
                                    'If WR call ClearMovesSub
                                    If ActivePiece.Range("A1").Value = "WR" _
                                        Then
                                            Set WhiteRook.RookMoves = Moves
                                            Call ClearMoves(ActivePiece, WhiteRook.RookMoves)
                                        Else
                                    End If
                                    'if WKN call ClearMovesSub
                                    If ActivePiece.Range("A1").Value = "WKN" _
                                        Then
                                            Set WhiteKnight.KnightMoves = Moves
                                            Call ClearMoves(ActivePiece, WhiteKnight.KnightMoves)
                                        Else
                                    End If
                                    'if WB call ClearMovesSub
                                    If ActivePiece.Range("A1").Value = "WB" _
                                        Then
                                            Set WhiteBishop.BishopMoves = Moves
                                            Call ClearMoves(ActivePiece, WhiteBishop.BishopMoves)
                                        Else
                                    End If
                                    'if WQ call ClearMovesSub
                                    If ActivePiece.Range("A1").Value = "WQ" _
                                        Then
                                            Set WhiteQueen.QueenMoves = Moves
                                            Call ClearMoves(ActivePiece, WhiteQueen.QueenMoves)
                                        Else
                                    End If
                                    'if WK call ClearMovesSub
                                    If ActivePiece.Range("A1").Value = "WK" _
                                        Then
                                            Set WhiteKing.KingMoves = Moves
                                            Set WhiteKing.ImpossibleMoves = GlobalImpMoves
                                            If WhiteKing.KingMoves Is Nothing And WhiteKing.ImpossibleMoves Is Nothing _
                                                Then
                                                    Set ClearRange = Nothing
                                                Else
                                                    If Not WhiteKing.KingMoves Is Nothing And WhiteKing.ImpossibleMoves Is Nothing _
                                                        Then
                                                            Set ClearRange = WhiteKing.KingMoves
                                                        Else
                                                            If WhiteKing.KingMoves Is Nothing And Not WhiteKing.ImpossibleMoves Is Nothing _
                                                                Then
                                                                    Set ClearRange = WhiteKing.ImpossibleMoves
                                                                Else
                                                                    Set ClearRange = Union(WhiteKing.KingMoves, WhiteKing.ImpossibleMoves)
                                                            End If
                                                    End If
                                            End If
                                            'If there are Castlemoves to add to the range clear
                                            Set WhiteKing.CastleMoves = GlobalCastleMoves
                                            If WhiteKing.CastleMoves Is Nothing _
                                                Then
                                                    'do nothing
                                                Else
                                                    If ClearRange Is Nothing _
                                                        Then
                                                            Set ClearRange = WhiteKing.CastleMoves
                                                        Else
                                                            Set ClearRange = Union(ClearRange, WhiteKing.CastleMoves)
                                                    End If
                                            End If
                                            Call ClearMoves(ActivePiece, ClearRange)
                                            If WhiteKingCheck = True Then ActivePiece.Interior.ColorIndex = 30
                                        Else
                                    End If
                                    'Clear ActivePiece and Moves variables
                                    Set ActivePiece = Nothing
                                    Set Moves = Nothing
                                Else
                            End If
                    End If
                    'If no Activepiece cell is present
                Else
                    ' If selection is 1 cell, value is not nothing, text length is 2-3 and first letter is "W"
                    If Target.Cells.Count = 1 And Target.Cells.Value <> "" And (Len(Target.Cells.Value) = 2 Or Len(Target.Cells.Value) = 3) _
                    And Left(Target.Cells.Value, 1) = "W" And Left(Target.Cells.Value, 2) <> "WK" _
                        Then
                            'Change color to yellow
                            Target.Interior.ColorIndex = 44
                        Else
                           ' If selection is 1 cell, value is not nothing, text length is 2-3 and first two letters "WK" for King
                            If Target.Cells.Count = 1 And Target.Cells.Value <> "" And (Len(Target.Cells.Value) = 2 Or Len(Target.Cells.Value) = 3) _
                            And Left(Target.Cells.Value, 2) = "WK" _
                                Then
                                    If Target.Interior.ColorIndex = 30 _
                                        Then
                                            'Set Colour of cell to yellow
                                            Target.Interior.ColorIndex = 44
                                        Else
                                            'Change color to yellow
                                            Target.Interior.ColorIndex = 44
                                    End If
                                Else
                                    'do nothing
                            End If
                    End If
                    'If White Pawn
                    If Target.Cells.Value = "WP" _
                        Then
                            Range("ActivePiece").Value = "Pawn"
                            WhitePawn.Player = "White"
                            Set WhitePawn.Target = Target
                            Call WhitePawn.CalcPawnMoves
                            Call WhitePawn.PaintMoves
                        Else
                            'If White Rook
                            If Target.Cells.Value = "WR" _
                                Then
                                    Range("ActivePiece").Value = "Rook"
                                    WhiteRook.Player = "White"
                                    Set WhiteRook.Target = Target
                                    Call WhiteRook.CalcRookMoves
                                    Call WhiteRook.PaintMoves
                                Else
                                    'If White Knight
                                    If Target.Cells.Value = "WKN" _
                                        Then
                                            Range("ActivePiece").Value = "Knight"
                                            WhiteKnight.Player = "White"
                                            Set WhiteKnight.Target = Target
                                            Call WhiteKnight.CalcKnightMoves
                                            Call WhiteKnight.PaintMoves
                                        Else
                                            'If White Bishop
                                            If Target.Cells.Value = "WB" _
                                                Then
                                                    Range("ActivePiece").Value = "Bishop"
                                                    WhiteBishop.Player = "White"
                                                    Set WhiteBishop.Target = Target
                                                    Call WhiteBishop.CalcBishopMoves
                                                    Call WhiteBishop.PaintMoves
                                                Else
                                                    'If White Queen
                                                    If Target.Cells.Value = "WQ" _
                                                        Then
                                                            Range("ActivePiece").Value = "Queen"
                                                            WhiteQueen.Player = "White"
                                                            Set WhiteQueen.Target = Target
                                                            Call WhiteQueen.CalcQueenMoves
                                                            Call WhiteQueen.PaintMoves
                                                        Else
                                                            'If White King
                                                            If Target.Cells.Value = "WK" _
                                                                Then
                                                                    'If White King in check
                                                                    If WhiteKingCheck = True _
                                                                        Then
                                                                            WhiteKing.Check = True
                                                                        Else
                                                                    End If
                                                                    Range("ActivePiece").Value = "King"
                                                                    WhiteKing.Player = "White"
                                                                    Set WhiteKing.Target = Target
                                                                    Call WhiteKing.CalcKingMoves
                                                                    Call WhiteKing.PaintMoves
                                                                    'If possible moves are none and King is in check call checkmate sub
                                                                    If WhiteKing.KingMoves Is Nothing And WhiteKingCheck = True Then Call Checkmate(WhiteKing.Player)
                                                                Else
                                                                    'Do nothing
                                                            End If
                                                    End If
                                            End If
                                    End If
                            End If
                    End If
                    'MsgBox "White"
            End If
            'MsgBox "White"
        Else
            'Methods if "ActivePlayer = Black"
            Player = "Black"
            'If a piece is selected
            If Not ActivePiece Is Nothing _
                Then
                    'If target is green ie is a potential movement square
                    If Target.Range("A1").Interior.ColorIndex = 50 Or Target.Interior.ColorIndex = 41 _
                        Then
                            'Cycle through Shapes find one in Target range
                            For Each SelShape In Sheet1.Shapes
                                ShpRow = SelShape.BottomRightCell.Row
                                ShpCol = SelShape.BottomRightCell.Column
                                 'If row and column match Target
                                 If ShpCol = Target.Column And ShpRow = Target.Row _
                                    Then
                                        Set TargetShape = SelShape
                                        'Move to taken pieces
                                        If TargetShape.Name <> "WP1" And TargetShape.Name <> "WP2" And TargetShape.Name <> "WP3" And TargetShape.Name <> "WP4" _
                                            And TargetShape.Name <> "WP5" And TargetShape.Name <> "WP6" And TargetShape.Name <> "WP7" And TargetShape.Name <> "WP8" _
                                            And TargetShape.Name <> "WR1" And TargetShape.Name <> "WR2" And TargetShape.Name <> "WKN1" And TargetShape.Name <> "WKN2" _
                                            And TargetShape.Name <> "WB1" And TargetShape.Name <> "WB2" And TargetShape.Name <> "WQ" _
                                            Then
                                                'Move to Pawn Promotion taken
                                                Call MoveToPPTaken(TargetShape)
                                            Else
                                                'Move to taken pieces
                                                Call MoveToTaken(TargetShape)
                                        End If
                                    Else
                                        'do nothing
                                 End If
                            Next
                            'Cycle through Shapes find one in ActivePiece range
                            For Each SelShape In Sheet1.Shapes
                                 ShpRow = SelShape.BottomRightCell.Row
                                 ShpCol = SelShape.BottomRightCell.Column
                                 'If row and column match
                                 If ShpCol = ActivePiece.Column And ShpRow = ActivePiece.Row _
                                    Then
                                        Set ActiveShape = SelShape
                                        'Move to Target range
                                        Call CenterMe(ActiveShape, Target)
                                    Else
                                        'Do nothing
                                 End If
                            Next
                        
                            'Change target value to Activepiece value > moving piece
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            '######################################################
                            
                            Target.Range("A1").Value = ActivePiece.Range("A1").Value
                            'MsgBox "Black"
                            
                            'If active piece is "BP" call ClearMoves Sub
                            If ActivePiece.Range("A1").Value = "BP" _
                                Then
                                    'If checked piece is not nothing
                                    If Not CheckedPiece Is Nothing _
                                        Then
                                            'If checked piece is Black King
                                            If CheckedPiece.Value = "BK" _
                                                Then
                                                    'Calc Black King possible moves
                                                    Set BlackKing.Target = CheckedPiece
                                                    BlackKing.Player = "Black"
                                                    BlackKing.CalcKingMoves
                                                    BlackKing.EvaluateCheck
                                                    'If Black King is in Check
                                                    If BlackKing.Check = True _
                                                        Then
                                                            'If no moves and in check disallow friendly piece movement
                                                            'Clear Target Move range
                                                            Target.Value = TargetValue
                                                            Call CenterMe(ActiveShape, ActivePiece2)
                                                            If TargetShape Is Nothing _
                                                                Then
                                                                    'Do nothing
                                                                Else
                                                                    Call CenterMe(TargetShape, Target)
                                                            End If
                                                            'Call Clearmoves using ActivePiece and whole of Chessboard for potential moves
                                                            Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                            'Ensure a cell can be reselected
                                                            Range("A1").Select
                                                            'Display message to Player
                                                            Message1 = MsgBox("Move does not break check", vbInformation + vbOKOnly, "Excel VBA Chess")
                                                            
                                                            'Reset ActivePiece and Moves
                                                            Set ActivePiece = Nothing
                                                            Set ActivePiece2 = Nothing
                                                            Set Moves = Nothing
                                                            'Ensure Events are re-enabled
                                                            Application.EnableEvents = True
                                                            Application.ScreenUpdating = True
                                                            Exit Sub
                                                        Else
                                                            'If Black King can move or not in Check anymore allow friendly piece movement
                                                            Application.ScreenUpdating = False
                                                            BlackPawn.Player = "Black"
                                                            Set BlackPawn.Target = ActivePiece2
                                                            'Turn target to nothing then switch back to "BP"
                                                            Target.Value = ""
                                                            BlackPawn.CalcPawnMoves
                                                            Target.Value = "BP"
                                                            Application.ScreenUpdating = True
                                                            'Call ClearMoves
                                                            Call ClearMoves(ActivePiece2, BlackPawn.PawnMoves)
                                                            'Call EvaluateCheckMoves
                                                            Call ClearMoves(CheckedPiece, Range("chessboard"))
                                                            'Set CheckedPiece to nothing
                                                            Set CheckedPiece = Nothing
                                                            BlackPawn.Player = "Black"
                                                            Set BlackPawn.Target = Target
                                                            Call BlackPawn.EvaluateCheckMoves
                                                            If WhiteKingCheck = True _
                                                                Then
                                                                    WhiteKing.Check = True
                                                                Else
                                                            End If
                                                            'Check if BP promoted
                                                            If Target.Row = 10 _
                                                                Then
                                                                    'If Pawn promoted
                                                                    PawnPromotion.Show
                                                                    'Cycle through Shapes find one in Target range
                                                                    For Each SelShape In Sheet1.Shapes
                                                                         ShpRow = SelShape.BottomRightCell.Row
                                                                         ShpCol = SelShape.BottomRightCell.Column
                                                                         'If row and column match
                                                                         If ShpCol = Target.Column And ShpRow = Target.Row _
                                                                            Then
                                                                                Set ActiveShape = SelShape
                                                                                'Move to taken pieces range
                                                                                
                                                                                'Move to taken pieces
                                                                                Call MoveToTaken(ActiveShape)
                                                                                
                                                                            Else
                                                                                'do nothing
                                                                         End If
                                                                    Next
                                                                    Select Case PromotedPiece
                                                                        Case "Queen"
                                                                            'Initialise Queen value, copy shape and increment
                                                                            Target.Value = "BQ"
                                                                            Set DuplicateShape = Sheet1.Shapes("BQ").Duplicate
                                                                            ShapeNumber = Sheet1.Shapes.Count
                                                                            Sheet1.Shapes(ShapeNumber).Name = "BQ" & BQIncrement
                                                                            Call CenterMe(Sheet1.Shapes(ShapeNumber), Target)
                                                                            Range("A1").Select
                                                                            BQIncrement = BQIncrement + 1
                                                                            'See if Black King in check
                                                                            BlackQueen.Player = "Black"
                                                                            Set BlackQueen.Target = Target
                                                                            Call BlackQueen.EvaluateCheckMoves
                                                                            If WhiteKingCheck = True _
                                                                                Then
                                                                                    WhiteKing.Check = True
                                                                                Else
                                                                            End If
                                                                            Set KingPosition = Nothing
                                                                        Case "Bishop"
                                                                            'Initialise Bishop value, copy shape and increment
                                                                            Target.Value = "BB"
                                                                            Set DuplicateShape = Sheet1.Shapes("BB1").Duplicate
                                                                            ShapeNumber = Sheet1.Shapes.Count
                                                                            Sheet1.Shapes(ShapeNumber).Name = "BB" & BBIncrement
                                                                            Call CenterMe(Sheet1.Shapes(ShapeNumber), Target)
                                                                            Range("A1").Select
                                                                            BBIncrement = BBIncrement + 1
                                                                            'See if Black King in check
                                                                            BlackBishop.Player = "White"
                                                                            Set BlackBishop.Target = Target
                                                                            Call BlackBishop.EvaluateCheckMoves
                                                                            If WhiteKingCheck = True _
                                                                                Then
                                                                                    WhiteKing.Check = True
                                                                                Else
                                                                            End If
                                                                            Set KingPosition = Nothing
                                                                        Case "Knight"
                                                                            'Initialise Knight value, copy shape and increment
                                                                            Target.Value = "BKN"
                                                                            Set DuplicateShape = Sheet1.Shapes("BKN1").Duplicate
                                                                            ShapeNumber = Sheet1.Shapes.Count
                                                                            Sheet1.Shapes(ShapeNumber).Name = "BKN" & BKNIncrement
                                                                            Call CenterMe(Sheet1.Shapes(ShapeNumber), Target)
                                                                            Range("A1").Select
                                                                            BKNIncrement = BKNIncrement + 1
                                                                            'See if Black King in check
                                                                            BlackKnight.Player = "black"
                                                                            Set BlackKnight.Target = Target
                                                                            Call BlackKnight.EvaluateCheckMoves
                                                                            If WhiteKingCheck = True _
                                                                                Then
                                                                                    WhiteKing.Check = True
                                                                                Else
                                                                            End If
                                                                            Set KingPosition = Nothing
                                                                        Case "Rook"
                                                                            'Initialise Knight value, copy shape and increment
                                                                            Target.Value = "BR"
                                                                            Set DuplicateShape = Sheet1.Shapes("BR1").Duplicate
                                                                            ShapeNumber = Sheet1.Shapes.Count
                                                                            Sheet1.Shapes(ShapeNumber).Name = "BR" & BRIncrement
                                                                            Call CenterMe(Sheet1.Shapes(ShapeNumber), Target)
                                                                            Range("A1").Select
                                                                            BRIncrement = BRIncrement + 1
                                                                            'See if Black King in check
                                                                            BlackRook.Player = "Black"
                                                                            Set BlackRook.Target = Target
                                                                            Call BlackRook.EvaluateCheckMoves
                                                                            If WhiteKingCheck = True _
                                                                                Then
                                                                                    WhiteKing.Check = True
                                                                                Else
                                                                            End If
                                                                            Set KingPosition = Nothing
                                                                    End Select
                                                                Else
                                                                    'Do nothing
                                                            End If
                                                    End If
                                                Else
                                            End If
                                        Else
                                            'If checked piece is nothing
                                            'Determine whether piece pinned
                                            Call FindKing("Black")
                                            Set BlackKing.Target = KingPosition
                                            
                                            BlackKing.Player = "Black"
                                            ActivePiece2.Value = ""
                                            BlackKing.CalcKingMoves
                                            BlackKing.EvaluateCheck
                                            ActivePiece2.Value = "BP"
                                            'If Black King is in Check
                                            If BlackKing.Check = True _
                                                Then
                                                    'If no moves and in check disallow friendly piece movement
                                                    'Clear Target Move range
                                                    Target.Value = TargetValue
                                                    Call CenterMe(ActiveShape, ActivePiece2)
                                                    If TargetShape Is Nothing _
                                                        Then
                                                            'Do nothing
                                                        Else
                                                            Call CenterMe(TargetShape, Target)
                                                    End If
                                                    'Call Clearmoves using ActivePiece and whole of Chessboard for potential moves
                                                    Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                    Call ClearMoves(KingPosition, Range("chessboard"))
                                                    'Ensure a cell can be reselected
                                                    Range("A1").Select
                                                    'Display message to Player
                                                    Message1 = MsgBox("Piece is in an absolute pin and cannot move here", vbInformation + vbOKOnly, "Excel VBA Chess")
                                                    
                                                    'Reset ActivePiece and Moves
                                                    Set ActivePiece = Nothing
                                                    Set ActivePiece2 = Nothing
                                                    Set Moves = Nothing
                                                    Set KingPosition = Nothing
                                                    'Ensure Events are re-enabled
                                                    Application.EnableEvents = True
                                                    Application.ScreenUpdating = True
                                                    Exit Sub
                                                Else
                                                    'Check for En Passant move from left
                                                    If ActivePiece2.Row = 4 And Target.Row = 6 And Target.Offset(0, -1).Value = "WP" _
                                                        Then
                                                            EnPassantMsg = MsgBox("White Player: Take Pawn 'en passant' from left side of Black Player?", vbYesNo + vbQuestion, "Excel VBA Chess")
                                                            If EnPassantMsg = vbYes _
                                                                Then
                                                                    Flag = "EP"
                                                                    For Each SelShape In Sheet1.Shapes
                                                                        ShpRow = SelShape.BottomRightCell.Row
                                                                        ShpCol = SelShape.BottomRightCell.Column
                                                                         'If row and column match Target
                                                                         If ShpCol = Target.Column And ShpRow = Target.Row _
                                                                            Then
                                                                                Set TargetShape = SelShape
                                                                                
                                                                                If TargetShape.Name <> "BP1" And TargetShape.Name <> "BP2" And TargetShape.Name <> "BP3" And TargetShape.Name <> "BP4" _
                                                                                    And TargetShape.Name <> "BP5" And TargetShape.Name <> "BP6" And TargetShape.Name <> "BP7" And TargetShape.Name <> "BP8" _
                                                                                    And TargetShape.Name <> "BR1" And TargetShape.Name <> "BR2" And TargetShape.Name <> "BKN1" And TargetShape.Name <> "BKN2" _
                                                                                    And TargetShape.Name <> "BB1" And TargetShape.Name <> "BB2" And TargetShape.Name <> "BQ" _
                                                                                    Then
                                                                                        'Move to Pawn Promotion taken
                                                                                        Call MoveToPPTaken(TargetShape)
                                                                                    Else
                                                                                        'Move to taken pieces
                                                                                        Call MoveToTaken(TargetShape)
                                                                                End If
                                                                            Else
                                                                                'do nothing
                                                                         End If
                                                                    Next
                                                                    Target.Value = ""
                                                                    'Cycle through Shapes find one in Target range
                                                                    For Each SelShape In Sheet1.Shapes
                                                                        ShpRow = SelShape.BottomRightCell.Row
                                                                        ShpCol = SelShape.BottomRightCell.Column
                                                                         'If row and column match Target
                                                                         If ShpCol = Target.Offset(0, -1).Column And ShpRow = Target.Offset(0, -1).Row _
                                                                            Then
                                                                                Call CenterMe(SelShape, Target.Offset(-1, 0))
                                                                            Else
                                                                                'do nothing
                                                                         End If
                                                                         
                                                                    Next
                                                                    
                                                                    'Call Algebraic Turns
                                                                    Call AlgebraicTurns(Target, TargetShape, ActivePiece, ActivePiece2, Flag)
                                                                    
                                                                    'Move pieces
                                                                    Target.Offset(0, -1).Value = ""
                                                                    Target.Offset(-1, 0).Value = "WP"
                                                                    ActivePiece2.Value = ""
                                                                    Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                                    
                                                                    'Reset ActivePiece and Moves
                                                                    Set ActivePiece = Nothing
                                                                    Set ActivePiece2 = Nothing
                                                                    Set Moves = Nothing
                                                                    Set KingPosition = Nothing
                                                                    
                                                                    Range("ActivePlayer").Value = "Black"
                                                                    
                                                                    Application.ScreenUpdating = True
                                                                    Application.EnableEvents = True
                                                                    
                                                                    Exit Sub
                                                                Else
                                                                    'Check for en passant move from right
                                                                    If ActivePiece2.Row = 4 And Target.Row = 6 And Target.Offset(0, 1).Value = "WP" _
                                                                        Then
                                                                            EnPassantMsg = MsgBox("White Player: Take Pawn 'en passant' from right side of Black Player?", vbYesNo + vbQuestion, "Excel VBA Chess")
                                                                            If EnPassantMsg = vbYes _
                                                                                Then
                                                                                    Flag = "EP"
                                                                                    For Each SelShape In Sheet1.Shapes
                                                                                        ShpRow = SelShape.BottomRightCell.Row
                                                                                        ShpCol = SelShape.BottomRightCell.Column
                                                                                         'If row and column match Target
                                                                                         If ShpCol = Target.Column And ShpRow = Target.Row _
                                                                                            Then
                                                                                                Set TargetShape = SelShape
                                                                                                
                                                                                                If TargetShape.Name <> "BP1" And TargetShape.Name <> "BP2" And TargetShape.Name <> "BP3" And TargetShape.Name <> "BP4" _
                                                                                                    And TargetShape.Name <> "BP5" And TargetShape.Name <> "BP6" And TargetShape.Name <> "BP7" And TargetShape.Name <> "BP8" _
                                                                                                    And TargetShape.Name <> "BR1" And TargetShape.Name <> "BR2" And TargetShape.Name <> "BKN1" And TargetShape.Name <> "BKN2" _
                                                                                                    And TargetShape.Name <> "BB1" And TargetShape.Name <> "BB2" And TargetShape.Name <> "BQ" _
                                                                                                    Then
                                                                                                        'Move to Pawn Promotion taken
                                                                                                        Call MoveToPPTaken(TargetShape)
                                                                                                    Else
                                                                                                        'Move to taken pieces
                                                                                                        Call MoveToTaken(TargetShape)
                                                                                                End If
                                                                                            Else
                                                                                                'do nothing
                                                                                         End If
                                                                                    Next
                                                                                    Target.Value = ""
                                                                                    'Cycle through Shapes find one in Target range
                                                                                    For Each SelShape In Sheet1.Shapes
                                                                                        ShpRow = SelShape.BottomRightCell.Row
                                                                                        ShpCol = SelShape.BottomRightCell.Column
                                                                                         'If row and column match Target
                                                                                         If ShpCol = Target.Offset(0, 1).Column And ShpRow = Target.Offset(0, 1).Row _
                                                                                            Then
                                                                                                Call CenterMe(SelShape, Target.Offset(-1, 0))
                                                                                            Else
                                                                                                'Do nothing
                                                                                         End If
                                                                                         
                                                                                    Next
                                                                                    
                                                                                    'Call Algebraic Turns
                                                                                    Call AlgebraicTurns(Target, TargetShape, ActivePiece, ActivePiece2, Flag)
                                                                                    
                                                                                    'Move pieces
                                                                                    Target.Offset(0, 1).Value = ""
                                                                                    Target.Offset(-1, 0).Value = "WP"
                                                                                    ActivePiece2.Value = ""
                                                                                    Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                                                    
                                                                                    'Reset ActivePiece and Moves
                                                                                    Set ActivePiece = Nothing
                                                                                    Set ActivePiece2 = Nothing
                                                                                    Set Moves = Nothing
                                                                                    Set KingPosition = Nothing
                                                                                    
                                                                                    Range("ActivePlayer").Value = "Black"
                                                                                    
                                                                                    Application.ScreenUpdating = True
                                                                                    Application.EnableEvents = True
                                                                                    
                                                                                    Exit Sub
                                                                                Else
                                                                                    'Do nothing
                                                                            End If
                                                                        Else
                                                                    End If
                                                            End If
                                                        Else
                                                            'Check for en passant move from right
                                                            If ActivePiece2.Row = 4 And Target.Row = 6 And Target.Offset(0, 1).Value = "WP" _
                                                                Then
                                                                    EnPassantMsg = MsgBox("White Player: Take Pawn 'en passant' from right side of Black Player?", vbYesNo + vbQuestion, "Excel VBA Chess")
                                                                    If EnPassantMsg = vbYes _
                                                                        Then
                                                                            Flag = "EP"
                                                                            For Each SelShape In Sheet1.Shapes
                                                                                ShpRow = SelShape.BottomRightCell.Row
                                                                                ShpCol = SelShape.BottomRightCell.Column
                                                                                 'If row and column match Target
                                                                                 If ShpCol = Target.Column And ShpRow = Target.Row _
                                                                                    Then
                                                                                        Set TargetShape = SelShape
                                                                                        
                                                                                        If TargetShape.Name <> "BP1" And TargetShape.Name <> "BP2" And TargetShape.Name <> "BP3" And TargetShape.Name <> "BP4" _
                                                                                            And TargetShape.Name <> "BP5" And TargetShape.Name <> "BP6" And TargetShape.Name <> "BP7" And TargetShape.Name <> "BP8" _
                                                                                            And TargetShape.Name <> "BR1" And TargetShape.Name <> "BR2" And TargetShape.Name <> "BKN1" And TargetShape.Name <> "BKN2" _
                                                                                            And TargetShape.Name <> "BB1" And TargetShape.Name <> "BB2" And TargetShape.Name <> "BQ" _
                                                                                            Then
                                                                                                'Move to Pawn Promotion taken
                                                                                                Call MoveToPPTaken(TargetShape)
                                                                                            Else
                                                                                                'Move to taken pieces
                                                                                                Call MoveToTaken(TargetShape)
                                                                                        End If
                                                                                    Else
                                                                                        'do nothing
                                                                                 End If
                                                                            Next
                                                                            Target.Value = ""
                                                                            'Cycle through Shapes find one in Target range
                                                                            For Each SelShape In Sheet1.Shapes
                                                                                ShpRow = SelShape.BottomRightCell.Row
                                                                                ShpCol = SelShape.BottomRightCell.Column
                                                                                 'If row and column match Target
                                                                                 If ShpCol = Target.Offset(0, 1).Column And ShpRow = Target.Offset(0, 1).Row _
                                                                                    Then
                                                                                        Call CenterMe(SelShape, Target.Offset(-1, 0))
                                                                                    Else
                                                                                        'Do nothing
                                                                                 End If
                                                                                 
                                                                            Next
                                                                            
                                                                            'Call Algebraic Turns
                                                                            Call AlgebraicTurns(Target, TargetShape, ActivePiece, ActivePiece2, Flag)
                                                                            
                                                                            'Move pieces
                                                                            Target.Offset(0, 1).Value = ""
                                                                            Target.Offset(-1, 0).Value = "WP"
                                                                            ActivePiece2.Value = ""
                                                                            Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                                            
                                                                            'Reset ActivePiece and Moves
                                                                            Set ActivePiece = Nothing
                                                                            Set ActivePiece2 = Nothing
                                                                            Set Moves = Nothing
                                                                            Set KingPosition = Nothing
                                                                            
                                                                            Range("ActivePlayer").Value = "Black"
                                                                            
                                                                            Application.ScreenUpdating = True
                                                                            Application.EnableEvents = True
                                                                            
                                                                            Exit Sub
                                                                        Else
                                                                            'Do nothing
                                                                    End If
                                                                Else
                                                            End If
                                                    End If
                                                    'Call ClearMoves
                                                    Set BlackPawn.PawnMoves = Moves
                                                    Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                    'Call EvaluateCheckMoves
                                                    BlackPawn.Player = "Black"
                                                    Set BlackPawn.Target = Target
                                                    Call BlackPawn.EvaluateCheckMoves
                                                    If WhiteKingCheck = True _
                                                        Then
                                                            WhiteKing.Check = True
                                                        Else
                                                    End If
                                                    Set KingPosition = Nothing
                                                    'Check if BP promoted
                                                    If Target.Row = 10 _
                                                        Then
                                                            'If Pawn promoted
                                                            PawnPromotion.Show
                                                            'Cycle through Shapes find one in Target range
                                                            For Each SelShape In Sheet1.Shapes
                                                                 ShpRow = SelShape.BottomRightCell.Row
                                                                 ShpCol = SelShape.BottomRightCell.Column
                                                                 'If row and column match
                                                                 If ShpCol = Target.Column And ShpRow = Target.Row _
                                                                    Then
                                                                        Set ActiveShape = SelShape
                                                                        'Move to taken pieces range
                                                                       
                                                                        'Move to taken pieces
                                                                        Call MoveToTaken(ActiveShape)
                                                                    Else
                                                                        'do nothing
                                                                 End If
                                                            Next
                                                            Select Case PromotedPiece
                                                                Case "Queen"
                                                                    'Initialise Queen value, copy shape and increment
                                                                    Target.Value = "BQ"
                                                                    Set DuplicateShape = Sheet1.Shapes("BQ").Duplicate
                                                                    ShapeNumber = Sheet1.Shapes.Count
                                                                    Sheet1.Shapes(ShapeNumber).Name = "BQ" & BQIncrement
                                                                    Call CenterMe(Sheet1.Shapes(ShapeNumber), Target)
                                                                    Range("A1").Select
                                                                    BQIncrement = BQIncrement + 1
                                                                    'See if Black King in check
                                                                    BlackQueen.Player = "Black"
                                                                    Set BlackQueen.Target = Target
                                                                    Call BlackQueen.EvaluateCheckMoves
                                                                    If WhiteKingCheck = True _
                                                                        Then
                                                                            WhiteKing.Check = True
                                                                        Else
                                                                    End If
                                                                    Set KingPosition = Nothing
                                                                Case "Bishop"
                                                                    'Initialise Bishop value, copy shape and increment
                                                                    Target.Value = "BB"
                                                                    Set DuplicateShape = Sheet1.Shapes("BB1").Duplicate
                                                                    ShapeNumber = Sheet1.Shapes.Count
                                                                    Sheet1.Shapes(ShapeNumber).Name = "BB" & BBIncrement
                                                                    Call CenterMe(Sheet1.Shapes(ShapeNumber), Target)
                                                                    Range("A1").Select
                                                                    BBIncrement = BBIncrement + 1
                                                                    'See if Black King in check
                                                                    BlackBishop.Player = "White"
                                                                    Set BlackBishop.Target = Target
                                                                    Call BlackBishop.EvaluateCheckMoves
                                                                    If WhiteKingCheck = True _
                                                                        Then
                                                                            WhiteKing.Check = True
                                                                        Else
                                                                    End If
                                                                    Set KingPosition = Nothing
                                                                Case "Knight"
                                                                    'Initialise Knight value, copy shape and increment
                                                                    Target.Value = "BKN"
                                                                    Set DuplicateShape = Sheet1.Shapes("BKN1").Duplicate
                                                                    ShapeNumber = Sheet1.Shapes.Count
                                                                    Sheet1.Shapes(ShapeNumber).Name = "BKN" & BKNIncrement
                                                                    Call CenterMe(Sheet1.Shapes(ShapeNumber), Target)
                                                                    Range("A1").Select
                                                                    BKNIncrement = BKNIncrement + 1
                                                                    'See if Black King in check
                                                                    BlackKnight.Player = "Black"
                                                                    Set BlackKnight.Target = Target
                                                                    Call BlackKnight.EvaluateCheckMoves
                                                                    If WhiteKingCheck = True _
                                                                        Then
                                                                            WhiteKing.Check = True
                                                                        Else
                                                                    End If
                                                                    Set KingPosition = Nothing
                                                                Case "Rook"
                                                                    'Initialise Knight value, copy shape and increment
                                                                    Target.Value = "BR"
                                                                    Set DuplicateShape = Sheet1.Shapes("BR1").Duplicate
                                                                    ShapeNumber = Sheet1.Shapes.Count
                                                                    Sheet1.Shapes(ShapeNumber).Name = "BR" & BRIncrement
                                                                    Call CenterMe(Sheet1.Shapes(ShapeNumber), Target)
                                                                    Range("A1").Select
                                                                    BRIncrement = BRIncrement + 1
                                                                    'See if Black King in check
                                                                    BlackRook.Player = "Black"
                                                                    Set BlackRook.Target = Target
                                                                    Call BlackRook.EvaluateCheckMoves
                                                                    If WhiteKingCheck = True _
                                                                        Then
                                                                            WhiteKing.Check = True
                                                                        Else
                                                                    End If
                                                                    Set KingPosition = Nothing
                                                            End Select
                                                        Else
                                                            'Do nothing
                                                    End If
                                                    
                                            End If
                                    End If
                                Else
                            End If
                            'If active piece is "BR" call ClearMoves Sub
                            If ActivePiece.Range("A1").Value = "BR" _
                                Then
                                    'If checked piece is not nothing
                                    If Not CheckedPiece Is Nothing _
                                        Then
                                            'If checked piece is Black King
                                            If CheckedPiece.Value = "BK" _
                                                Then
                                                    'Calc Black King possible moves
                                                    Set BlackKing.Target = CheckedPiece
                                                    BlackKing.Player = "Black"
                                                    BlackKing.CalcKingMoves
                                                    BlackKing.EvaluateCheck
                                                    'If Black King is in Check
                                                    If BlackKing.Check = True _
                                                        Then
                                                            'If no moves and in check disallow friendly piece movement
                                                            'Clear Target Move range
                                                            Target.Value = TargetValue
                                                            Call CenterMe(ActiveShape, ActivePiece2)
                                                            If TargetShape Is Nothing _
                                                                Then
                                                                    'Do nothing
                                                                Else
                                                                    Call CenterMe(TargetShape, Target)
                                                            End If
                                                            'Call Clearmoves using ActivePiece and whole of Chessboard for potential moves
                                                            Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                            'Ensure a cell can be reselected
                                                            Range("A1").Select
                                                            'Display message to Player
                                                            Message1 = MsgBox("Move does not break check", vbInformation + vbOKOnly, "Excel VBA Chess")
                                                            
                                                            'Reset ActivePiece and Moves
                                                            Set ActivePiece = Nothing
                                                            Set ActivePiece2 = Nothing
                                                            Set Moves = Nothing
                                                            'Ensure Events are re-enabled
                                                            Application.EnableEvents = True
                                                            Application.ScreenUpdating = True
                                                            Exit Sub
                                                        Else
                                                            'If Black King can move or not in Check anymore allow friendly piece movement
                                                            Application.ScreenUpdating = False
                                                            BlackRook.Player = "Black"
                                                            Set BlackRook.Target = ActivePiece2
                                                            'Turn target to nothing then switch back to "BR"
                                                            Target.Value = ""
                                                            BlackRook.CalcRookMoves
                                                            Target.Value = "BR"
                                                            Application.ScreenUpdating = True
                                                            'Call ClearMoves
                                                            Call ClearMoves(ActivePiece2, BlackRook.RookMoves)
                                                            'Call EvaluateCheckMoves
                                                            Call ClearMoves(CheckedPiece, Range("chessboard"))
                                                            'Set CheckedPiece to nothing
                                                            Set CheckedPiece = Nothing
                                                            BlackRook.Player = "Black"
                                                            Set BlackRook.Target = Target
                                                            Call BlackRook.EvaluateCheckMoves
                                                            If WhiteKingCheck = True _
                                                                Then
                                                                    WhiteKing.Check = True
                                                                Else
                                                            End If
                                                    End If
                                                Else
                                            End If
                                        Else
                                            'If checked piece is nothing
                                            'Determine whether piece pinned
                                            Call FindKing("Black")
                                            Set BlackKing.Target = KingPosition
                                            
                                            BlackKing.Player = "Black"
                                            ActivePiece2.Value = ""
                                            BlackKing.CalcKingMoves
                                            BlackKing.EvaluateCheck
                                            ActivePiece2.Value = "BR"
                                            'If Black King is in Check
                                            If BlackKing.Check = True _
                                                Then
                                                    'If no moves and in check disallow friendly piece movement
                                                    'Clear Target Move range
                                                    Target.Value = TargetValue
                                                    Call CenterMe(ActiveShape, ActivePiece2)
                                                    If TargetShape Is Nothing _
                                                        Then
                                                            'Do nothing
                                                        Else
                                                            Call CenterMe(TargetShape, Target)
                                                    End If
                                                    'Call Clearmoves using ActivePiece and whole of Chessboard for potential moves
                                                    Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                    Call ClearMoves(KingPosition, Range("chessboard"))
                                                    'Ensure a cell can be reselected
                                                    Range("A1").Select
                                                    'Display message to Player
                                                    Message1 = MsgBox("Piece is in an absolute pin and cannot move here", vbInformation + vbOKOnly, "Excel VBA Chess")
                                                    
                                                    'Reset ActivePiece and Moves
                                                    Set ActivePiece = Nothing
                                                    Set ActivePiece2 = Nothing
                                                    Set Moves = Nothing
                                                    Set KingPosition = Nothing
                                                    'Ensure Events are re-enabled
                                                    Application.EnableEvents = True
                                                    Application.ScreenUpdating = True
                                                    Exit Sub
                                                Else
                                                    'Call ClearMoves
                                                    Set BlackRook.RookMoves = Moves
                                                    Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                    'Call EvaluateCheckMoves
                                                    BlackRook.Player = "Black"
                                                    Set BlackRook.Target = Target
                                                    Call BlackRook.EvaluateCheckMoves
                                                    If WhiteKingCheck = True _
                                                        Then
                                                            WhiteKing.Check = True
                                                        Else
                                                    End If
                                                    Set KingPosition = Nothing
                                            End If
                                    End If
                                Else
                            End If
                            'If active piece is "BKN" call ClearMoves Sub
                            If ActivePiece.Range("A1").Value = "BKN" _
                                Then
                                    'If checked piece is not nothing
                                    If Not CheckedPiece Is Nothing _
                                        Then
                                            'If checked piece is Black King
                                            If CheckedPiece.Value = "BK" _
                                                Then
                                                    'Calc Black King possible moves
                                                    Set BlackKing.Target = CheckedPiece
                                                    BlackKing.Player = "Black"
                                                    BlackKing.CalcKingMoves
                                                    BlackKing.EvaluateCheck
                                                    'If Black King is in Check
                                                    If BlackKing.Check = True _
                                                        Then
                                                            'If no moves and in check disallow friendly piece movement
                                                            'Clear Target Move range
                                                            Target.Value = TargetValue
                                                            Call CenterMe(ActiveShape, ActivePiece2)
                                                            If TargetShape Is Nothing _
                                                                Then
                                                                    'Do nothing
                                                                Else
                                                                    Call CenterMe(TargetShape, Target)
                                                            End If
                                                            'Call Clearmoves using ActivePiece and whole of Chessboard for potential moves
                                                            Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                            'Ensure a cell can be reselected
                                                            Range("A1").Select
                                                            'Display message to Player
                                                            Message1 = MsgBox("Move does not break check", vbInformation + vbOKOnly, "Excel VBA Chess")
                                                            
                                                            'Reset ActivePiece and Moves
                                                            Set ActivePiece = Nothing
                                                            Set ActivePiece2 = Nothing
                                                            Set Moves = Nothing
                                                            'Ensure Events are re-enabled
                                                            Application.EnableEvents = True
                                                            Application.ScreenUpdating = True
                                                            Exit Sub
                                                        Else
                                                            'If Black King can move or not in Check anymore allow friendly piece movement
                                                            Application.ScreenUpdating = False
                                                            BlackKnight.Player = "Black"
                                                            Set BlackKnight.Target = ActivePiece2
                                                            'Turn target to nothing then switch back to "BKN"
                                                            Target.Value = ""
                                                            BlackKnight.CalcKnightMoves
                                                            Target.Value = "BKN"
                                                            Application.ScreenUpdating = True
                                                            'Call ClearMoves
                                                            Call ClearMoves(ActivePiece2, BlackKnight.KnightMoves)
                                                            'Call EvaluateCheckMoves
                                                            Call ClearMoves(CheckedPiece, Range("chessboard"))
                                                            'Set CheckedPiece to nothing
                                                            Set CheckedPiece = Nothing
                                                            BlackKnight.Player = "Black"
                                                            Set BlackKnight.Target = Target
                                                            Call BlackKnight.EvaluateCheckMoves
                                                            If WhiteKingCheck = True _
                                                                Then
                                                                    WhiteKing.Check = True
                                                                Else
                                                            End If
                                                    End If
                                                Else
                                            End If
                                        Else
                                            'If checked piece is nothing
                                            'Determine whether piece pinned
                                            Call FindKing("Black")
                                            Set BlackKing.Target = KingPosition
                                            
                                            BlackKing.Player = "Black"
                                            ActivePiece2.Value = ""
                                            BlackKing.CalcKingMoves
                                            BlackKing.EvaluateCheck
                                            ActivePiece2.Value = "BKN"
                                            'If Black King is in Check
                                            If BlackKing.Check = True _
                                                Then
                                                    'If no moves and in check disallow friendly piece movement
                                                    'Clear Target Move range
                                                    Target.Value = TargetValue
                                                    Call CenterMe(ActiveShape, ActivePiece2)
                                                    If TargetShape Is Nothing _
                                                        Then
                                                            'Do nothing
                                                        Else
                                                            Call CenterMe(TargetShape, Target)
                                                    End If
                                                    'Call Clearmoves using ActivePiece and whole of Chessboard for potential moves
                                                    Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                    Call ClearMoves(KingPosition, Range("chessboard"))
                                                    'Ensure a cell can be reselected
                                                    Range("A1").Select
                                                    'Display message to Player
                                                    Message1 = MsgBox("Piece is in an absolute pin and cannot move here", vbInformation + vbOKOnly, "Excel VBA Chess")
                                                    
                                                    'Reset ActivePiece and Moves
                                                    Set ActivePiece = Nothing
                                                    Set ActivePiece2 = Nothing
                                                    Set Moves = Nothing
                                                    Set KingPosition = Nothing
                                                    'Ensure Events are re-enabled
                                                    Application.EnableEvents = True
                                                    Application.ScreenUpdating = True
                                                    Exit Sub
                                                Else
                                                    'Call ClearMoves
                                                    Set BlackKnight.KnightMoves = Moves
                                                    Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                    'Call EvaluateCheckMoves
                                                    BlackKnight.Player = "Black"
                                                    Set BlackKnight.Target = Target
                                                    Call BlackKnight.EvaluateCheckMoves
                                                    If WhiteKingCheck = True _
                                                        Then
                                                            WhiteKing.Check = True
                                                        Else
                                                    End If
                                                    Set KingPosition = Nothing
                                            End If
                                    End If
                                Else
                            End If
                            'If active piece is "BB" call ClearMoves Sub
                            If ActivePiece.Range("A1").Value = "BB" _
                                Then
                                    'If checked piece is not nothing
                                    If Not CheckedPiece Is Nothing _
                                        Then
                                            'If checked piece is Black King
                                            If CheckedPiece.Value = "BK" _
                                                Then
                                                    'Calc Black King possible moves
                                                    Set BlackKing.Target = CheckedPiece
                                                    BlackKing.Player = "Black"
                                                    BlackKing.CalcKingMoves
                                                    BlackKing.EvaluateCheck
                                                    'If Black King is in Check
                                                    If BlackKing.Check = True _
                                                        Then
                                                            'If no moves and in check disallow friendly piece movement
                                                            'Clear Target Move range
                                                            Target.Value = TargetValue
                                                            Call CenterMe(ActiveShape, ActivePiece2)
                                                            If TargetShape Is Nothing _
                                                                Then
                                                                    'Do nothing
                                                                Else
                                                                    Call CenterMe(TargetShape, Target)
                                                            End If
                                                            'Call Clearmoves using ActivePiece and whole of Chessboard for potential moves
                                                            Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                            'Ensure a cell can be reselected
                                                            Range("A1").Select
                                                            'Display message to Player
                                                            Message1 = MsgBox("Move does not break check", vbInformation + vbOKOnly, "Excel VBA Chess")
                                                            
                                                            'Reset ActivePiece and Moves
                                                            Set ActivePiece = Nothing
                                                            Set ActivePiece2 = Nothing
                                                            Set Moves = Nothing
                                                            'Ensure Events are re-enabled
                                                            Application.EnableEvents = True
                                                            Application.ScreenUpdating = True
                                                            Exit Sub
                                                        Else
                                                            'If Black King can move or not in Check anymore allow friendly piece movement
                                                            Application.ScreenUpdating = False
                                                            BlackBishop.Player = "Black"
                                                            Set BlackBishop.Target = ActivePiece2
                                                            'Turn target to nothing then switch back to "BB"
                                                            Target.Value = ""
                                                            BlackBishop.CalcBishopMoves
                                                            Target.Value = "BB"
                                                            Application.ScreenUpdating = True
                                                            'Call ClearMoves
                                                            Call ClearMoves(ActivePiece2, BlackBishop.BishopMoves)
                                                            'Call EvaluateCheckMoves
                                                            Call ClearMoves(CheckedPiece, Range("chessboard"))
                                                            'Set CheckedPiece to nothing
                                                            Set CheckedPiece = Nothing
                                                            BlackBishop.Player = "Black"
                                                            Set BlackBishop.Target = Target
                                                            Call BlackBishop.EvaluateCheckMoves
                                                            If WhiteKingCheck = True _
                                                                Then
                                                                    WhiteKing.Check = True
                                                                Else
                                                            End If
                                                    End If
                                                Else
                                            End If
                                        Else
                                            'If checked piece is nothing
                                            'Determine whether piece pinned
                                            Call FindKing("Black")
                                            Set BlackKing.Target = KingPosition
                                            
                                            BlackKing.Player = "Black"
                                            ActivePiece2.Value = ""
                                            BlackKing.CalcKingMoves
                                            BlackKing.EvaluateCheck
                                            ActivePiece2.Value = "BB"
                                            'If Black King is in Check
                                            If BlackKing.Check = True _
                                                Then
                                                    'If no moves and in check disallow friendly piece movement
                                                    'Clear Target Move range
                                                    Target.Value = TargetValue
                                                    Call CenterMe(ActiveShape, ActivePiece2)
                                                    If TargetShape Is Nothing _
                                                        Then
                                                            'Do nothing
                                                        Else
                                                            Call CenterMe(TargetShape, Target)
                                                    End If
                                                    'Call Clearmoves using ActivePiece and whole of Chessboard for potential moves
                                                    Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                    Call ClearMoves(KingPosition, Range("chessboard"))
                                                    'Ensure a cell can be reselected
                                                    Range("A1").Select
                                                    'Display message to Player
                                                    Message1 = MsgBox("Piece is in an absolute pin and cannot move here", vbInformation + vbOKOnly, "Excel VBA Chess")
                                                    
                                                    'Reset ActivePiece and Moves
                                                    Set ActivePiece = Nothing
                                                    Set ActivePiece2 = Nothing
                                                    Set Moves = Nothing
                                                    Set KingPosition = Nothing
                                                    'Ensure Events are re-enabled
                                                    Application.EnableEvents = True
                                                    Application.ScreenUpdating = True
                                                    Exit Sub
                                                Else
                                                    'Call ClearMoves
                                                    Set BlackBishop.BishopMoves = Moves
                                                    Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                    'Call EvaluateCheckMoves
                                                    BlackBishop.Player = "Black"
                                                    Set BlackBishop.Target = Target
                                                    Call BlackBishop.EvaluateCheckMoves
                                                    If WhiteKingCheck = True _
                                                        Then
                                                            WhiteKing.Check = True
                                                        Else
                                                    End If
                                                    Set KingPosition = Nothing
                                            End If
                                    End If
                                Else
                            End If
                            'If active piece is "BQ" call ClearMoves Sub
                            If ActivePiece.Range("A1").Value = "BQ" _
                                Then
                                    'If checked piece is not nothing
                                    If Not CheckedPiece Is Nothing _
                                        Then
                                            'If checked piece is Black King
                                            If CheckedPiece.Value = "BK" _
                                                Then
                                                    'Calc Black King possible moves
                                                    Set BlackKing.Target = CheckedPiece
                                                    BlackKing.Player = "Black"
                                                    BlackKing.CalcKingMoves
                                                    BlackKing.EvaluateCheck
                                                    'If Black King is in Check
                                                    If BlackKing.Check = True _
                                                        Then
                                                            'If no moves and in check disallow friendly piece movement
                                                            'Clear Target Move range
                                                            Target.Value = TargetValue
                                                            Call CenterMe(ActiveShape, ActivePiece2)
                                                            If TargetShape Is Nothing _
                                                                Then
                                                                    'Do nothing
                                                                Else
                                                                    Call CenterMe(TargetShape, Target)
                                                            End If
                                                            'Call Clearmoves using ActivePiece and whole of Chessboard for potential moves
                                                            Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                            'Ensure a cell can be reselected
                                                            Range("A1").Select
                                                            'Display message to Player
                                                            Message1 = MsgBox("Move does not break check", vbInformation + vbOKOnly, "Excel VBA Chess")
                                                            
                                                            'Reset ActivePiece and Moves
                                                            Set ActivePiece = Nothing
                                                            Set ActivePiece2 = Nothing
                                                            Set Moves = Nothing
                                                            'Ensure Events are re-enabled
                                                            Application.EnableEvents = True
                                                            Application.ScreenUpdating = True
                                                            Exit Sub
                                                        Else
                                                            'If Black King can move or not in Check anymore allow friendly piece movement
                                                            Application.ScreenUpdating = False
                                                            BlackQueen.Player = "Black"
                                                            Set BlackQueen.Target = ActivePiece2
                                                            'Turn target to nothing then switch back to "BQ"
                                                            Target.Value = ""
                                                            BlackQueen.CalcQueenMoves
                                                            Target.Value = "BQ"
                                                            Application.ScreenUpdating = True
                                                            'Call ClearMoves
                                                            Call ClearMoves(ActivePiece2, BlackQueen.QueenMoves)
                                                            'Call EvaluateCheckMoves
                                                            Call ClearMoves(CheckedPiece, Range("chessboard"))
                                                            'Set CheckedPiece to nothing
                                                            Set CheckedPiece = Nothing
                                                            BlackQueen.Player = "Black"
                                                            Set BlackQueen.Target = Target
                                                            Call BlackQueen.EvaluateCheckMoves
                                                            If WhiteKingCheck = True _
                                                                Then
                                                                    WhiteKing.Check = True
                                                                Else
                                                            End If
                                                    End If
                                                Else
                                            End If
                                        Else
                                            'If checked piece is nothing
                                            'Determine whether piece pinned
                                            Call FindKing("Black")
                                            Set BlackKing.Target = KingPosition
                                            
                                            BlackKing.Player = "Black"
                                            ActivePiece2.Value = ""
                                            BlackKing.CalcKingMoves
                                            BlackKing.EvaluateCheck
                                            ActivePiece2.Value = "BQ"
                                            'If Black King is in Check
                                            If BlackKing.Check = True _
                                                Then
                                                    'If no moves and in check disallow friendly piece movement
                                                    'Clear Target Move range
                                                    Target.Value = TargetValue
                                                    Call CenterMe(ActiveShape, ActivePiece2)
                                                    If TargetShape Is Nothing _
                                                        Then
                                                            'Do nothing
                                                        Else
                                                            Call CenterMe(TargetShape, Target)
                                                    End If
                                                    'Call Clearmoves using ActivePiece and whole of Chessboard for potential moves
                                                    Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                    Call ClearMoves(KingPosition, Range("chessboard"))
                                                    'Ensure a cell can be reselected
                                                    Range("A1").Select
                                                    'Display message to Player
                                                    Message1 = MsgBox("Piece is in an absolute pin and cannot move here", vbInformation + vbOKOnly, "Excel VBA Chess")
                                                    
                                                    'Reset ActivePiece and Moves
                                                    Set ActivePiece = Nothing
                                                    Set ActivePiece2 = Nothing
                                                    Set Moves = Nothing
                                                    Set KingPosition = Nothing
                                                    'Ensure Events are re-enabled
                                                    Application.EnableEvents = True
                                                    Application.ScreenUpdating = True
                                                    Exit Sub
                                                Else
                                                    'Call ClearMoves
                                                    Set BlackQueen.QueenMoves = Moves
                                                    Call ClearMoves(ActivePiece2, Range("chessboard"))
                                                    'Call EvaluateCheckMoves
                                                    BlackQueen.Player = "Black"
                                                    Set BlackQueen.Target = Target
                                                    Call BlackQueen.EvaluateCheckMoves
                                                    If WhiteKingCheck = True _
                                                        Then
                                                            WhiteKing.Check = True
                                                        Else
                                                    End If
                                                    Set KingPosition = Nothing
                                            End If
                                    End If
                                Else
                            End If
                            'If active piece is "BK" call ClearMoves Sub
                            If ActivePiece.Range("A1").Value = "BK" _
                                Then
                                    Set BlackKing.KingMoves = Moves
                                    Set BlackKing.ImpossibleMoves = GlobalImpMoves
                                    If BlackKing.KingMoves Is Nothing And BlackKing.ImpossibleMoves Is Nothing _
                                        Then
                                            Set ClearRange = Nothing
                                        Else
                                            If Not BlackKing.KingMoves Is Nothing And BlackKing.ImpossibleMoves Is Nothing _
                                                Then
                                                    Set ClearRange = BlackKing.KingMoves
                                                Else
                                                    If BlackKing.KingMoves Is Nothing And Not BlackKing.ImpossibleMoves Is Nothing _
                                                        Then
                                                            Set ClearRange = BlackKing.ImpossibleMoves
                                                        Else
                                                            Set ClearRange = Union(BlackKing.KingMoves, BlackKing.ImpossibleMoves)
                                                    End If
                                            End If
                                    End If
                                    'If there are Castlemoves to add to the range clear
                                    Set BlackKing.CastleMoves = GlobalCastleMoves
                                    If BlackKing.CastleMoves Is Nothing _
                                        Then
                                            'do nothing
                                        Else
                                            If ClearRange Is Nothing _
                                                Then
                                                    Set ClearRange = BlackKing.CastleMoves
                                                Else
                                                    Set ClearRange = Union(ClearRange, BlackKing.CastleMoves)
                                            End If
                                            'Move Rook to protect King
                                            If Target.Column = 8 _
                                                Then
                                                    ActivePiece2.Offset(0, 3).Value = ""
                                                    ActivePiece2.Offset(0, 1).Value = "BR"
                                                    Call CenterMe(Sheet1.Shapes("BR2"), Range("G3"))
                                                    Flag = "Kingside Castling"
                                                Else
                                                    If Target.Column = 4 _
                                                        Then
                                                            ActivePiece2.Offset(0, -4).Value = ""
                                                            ActivePiece2.Offset(0, -1).Value = "BR"
                                                            Call CenterMe(Sheet1.Shapes("BR1"), Range("E3"))
                                                            Flag = "Queenside Castling"
                                                        Else
                                                            'Do nothing
                                                    End If
                                            End If
                                    End If
                                    Call ClearMoves(ActivePiece, ClearRange)
                                    BlackKingCheck = False
                                    BlackKingMoved = True
                                    'Set CheckedPiece to nothing
                                    Set CheckedPiece = Nothing
                                Else
                            End If
                            
                            If (WhiteKingCheck = True Or BlackKingCheck = True) Then Flag = "Check"
                            
                            'Call Algebraic Turns
                            Call AlgebraicTurns(Target, TargetShape, ActivePiece, ActivePiece2, Flag)
                            
                            'Change activepiece2 value to "" > moving piece
                            ActivePiece2.Range("A1").Value = ""
                            Set ActivePiece = Nothing
                            Set ActivePiece2 = Nothing
                            Set Moves = Nothing
                            Call Increment(Range("Turn").Value)
                            'MsgBox "Done"
                            Call Send_Email_With_snapshot
                            
                            'Change Active Player to "White"
                            Range("ActivePlayer").Value = "White"
                        Else
                            'If target square is yellow ie is selected square > deselect
                            If Target.Range("A1").Interior.ColorIndex = 44 Or Target.Interior.ColorIndex = 30 _
                                Then
                                    'If BP Call ClearMovesSub
                                    If ActivePiece.Range("A1").Value = "BP" _
                                        Then
                                            Set BlackPawn.PawnMoves = Moves
                                            Call ClearMoves(ActivePiece, BlackPawn.PawnMoves)
                                        Else
                                    End If
                                    'If BR Call ClearMovesSub
                                    If ActivePiece.Range("A1").Value = "BR" _
                                        Then
                                            Set BlackRook.RookMoves = Moves
                                            Call ClearMoves(ActivePiece, BlackRook.RookMoves)
                                        Else
                                    End If
                                    'if BKN call ClearMovesSub
                                    If ActivePiece.Range("A1").Value = "BKN" _
                                        Then
                                            Set BlackKnight.KnightMoves = Moves
                                            Call ClearMoves(ActivePiece, BlackKnight.KnightMoves)
                                        Else
                                    End If
                                    'if BB call ClearMovesSub
                                    If ActivePiece.Range("A1").Value = "BB" _
                                        Then
                                            Set BlackBishop.BishopMoves = Moves
                                            Call ClearMoves(ActivePiece, BlackBishop.BishopMoves)
                                        Else
                                    End If
                                    'if BQ call ClearMovesSub
                                    If ActivePiece.Range("A1").Value = "BQ" _
                                        Then
                                            Set BlackQueen.QueenMoves = Moves
                                            Call ClearMoves(ActivePiece, BlackQueen.QueenMoves)
                                        Else
                                    End If
                                    'if BK call ClearMovesSub
                                    If ActivePiece.Range("A1").Value = "BK" _
                                        Then
                                            Set BlackKing.KingMoves = Moves
                                            Set BlackKing.ImpossibleMoves = GlobalImpMoves
                                            If BlackKing.KingMoves Is Nothing And BlackKing.ImpossibleMoves Is Nothing _
                                                Then
                                                    Set ClearRange = Nothing
                                                Else
                                                    If Not BlackKing.KingMoves Is Nothing And BlackKing.ImpossibleMoves Is Nothing _
                                                        Then
                                                            Set ClearRange = BlackKing.KingMoves
                                                        Else
                                                            If BlackKing.KingMoves Is Nothing And Not BlackKing.ImpossibleMoves Is Nothing _
                                                                Then
                                                                    Set ClearRange = BlackKing.ImpossibleMoves
                                                                Else
                                                                    Set ClearRange = Union(BlackKing.KingMoves, BlackKing.ImpossibleMoves)
                                                            End If
                                                    End If
                                            End If
                                            'If there are Castlemoves to add to the range clear
                                            Set BlackKing.CastleMoves = GlobalCastleMoves
                                            If BlackKing.CastleMoves Is Nothing _
                                                Then
                                                    'do nothing
                                                Else
                                                    If ClearRange Is Nothing _
                                                        Then
                                                            Set ClearRange = BlackKing.CastleMoves
                                                        Else
                                                            Set ClearRange = Union(ClearRange, BlackKing.CastleMoves)
                                                    End If
                                            End If
                                            Call ClearMoves(ActivePiece, ClearRange)
                                            If BlackKingCheck = True Then ActivePiece.Interior.ColorIndex = 30
                                        Else
                                    End If
                                    'Clear ActivePiece and Moves variables
                                    Set ActivePiece = Nothing
                                    Set Moves = Nothing
                                Else
                            End If
                    End If
                    'If no Activepiece cell is present
                Else
                    'If selection is 1 cell, value is not nothing, text length is 2-3 and first letter is "B"
                    If Target.Cells.Count = 1 And Target.Cells.Value <> "" And (Len(Target.Cells.Value) = 2 Or Len(Target.Cells.Value) = 3) _
                    And Left(Target.Cells.Value, 1) = "B" And Left(Target.Cells.Value, 2) <> "BK" _
                        Then
                            'Change color to yellow
                            Target.Interior.ColorIndex = 44
                        Else
                            ' If selection is 1 cell, value is not nothing, text length is 2-3 and first two letters "BK" for King
                            If Target.Cells.Count = 1 And Target.Cells.Value <> "" And (Len(Target.Cells.Value) = 2 Or Len(Target.Cells.Value) = 3) _
                            And Left(Target.Cells.Value, 2) = "BK" _
                                Then
                                    If Target.Interior.ColorIndex = 30 _
                                        Then
                                            'Set Colour of cell to yellow
                                            Target.Interior.ColorIndex = 44
                                        Else
                                            'Change color to yellow
                                            Target.Interior.ColorIndex = 44
                                    End If
                                Else
                                    'do nothing
                            End If
                    End If
                    'If Black Pawn
                    If Target.Cells.Value = "BP" _
                        Then
                            Range("ActivePiece").Value = "Pawn"
                            BlackPawn.Player = "Black"
                            Set BlackPawn.Target = Target
                            Call BlackPawn.CalcPawnMoves
                            Call BlackPawn.PaintMoves
                        Else
                            'If Black Rook
                            If Target.Cells.Value = "BR" _
                                Then
                                    Range("ActivePiece").Value = "Rook"
                                    BlackRook.Player = "Black"
                                    Set BlackRook.Target = Target
                                    Call BlackRook.CalcRookMoves
                                    Call BlackRook.PaintMoves
                                Else
                                    'If Black Knight
                                    If Target.Cells.Value = "BKN" _
                                        Then
                                            Range("ActivePiece").Value = "Knight"
                                            BlackKnight.Player = "Black"
                                            Set BlackKnight.Target = Target
                                            Call BlackKnight.CalcKnightMoves
                                            Call BlackKnight.PaintMoves
                                        Else
                                            'If Black Bishop
                                            If Target.Cells.Value = "BB" _
                                                Then
                                                    Range("ActivePiece").Value = "Bishop"
                                                    BlackBishop.Player = "Black"
                                                    Set BlackBishop.Target = Target
                                                    Call BlackBishop.CalcBishopMoves
                                                    Call BlackBishop.PaintMoves
                                                Else
                                                    'If Black Queen
                                                    If Target.Cells.Value = "BQ" _
                                                        Then
                                                            Range("ActivePiece").Value = "Queen"
                                                            BlackQueen.Player = "Black"
                                                            Set BlackQueen.Target = Target
                                                            Call BlackQueen.CalcQueenMoves
                                                            Call BlackQueen.PaintMoves
                                                        Else
                                                            'If Black King
                                                            If Target.Cells.Value = "BK" _
                                                                Then
                                                                    'If Black King in check
                                                                    If BlackKingCheck = True _
                                                                        Then
                                                                            BlackKing.Check = True
                                                                        Else
                                                                    End If
                                                                    Range("ActivePiece").Value = "King"
                                                                    BlackKing.Player = "Black"
                                                                    Set BlackKing.Target = Target
                                                                    Call BlackKing.CalcKingMoves
                                                                    Call BlackKing.PaintMoves
                                                                    'If possible moves are none and King is in check call checkmate sub
                                                                    If BlackKing.KingMoves Is Nothing And BlackKingCheck = True Then Call Checkmate(BlackKing.Player)
                                                                Else
                                                                    'Do nothing
                                                            End If
                                                    End If
                                            End If
                                    End If
                            End If
                    End If
            End If
            'MsgBox "CHECK"
    End If
    'MsgBox "Done"
    Set ActiveShape = Nothing
    Set TargetShape = Nothing
    Player = Empty
    PromotedPiece = Empty
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    'MsgBox "CHECK"

End Sub

Public Sub EnableEvents()

    Application.EnableEvents = True

End Sub

Public Sub DisableEvents()

    Application.EnableEvents = False

End Sub
