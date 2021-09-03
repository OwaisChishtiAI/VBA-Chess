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
    
    CreateWebQuery Sheet3.Range("A1"), url
End Sub
