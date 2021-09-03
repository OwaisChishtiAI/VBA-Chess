Option Explicit

Sub SelectAnEntireColumn()

    'Select column based on position
    'ActiveSheet.ListObjects("myTable").ListColumns(2).Range.Select

    'Select column based on name
    'Sheet5.ListObjects("myTable").ListColumns("Category").Range.Select
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Sheet3")
    Dim whitePlayer As String
    Dim blackPlayer As String

    'MsgBox Sheet5.Range("C8").Value
    Dim wr As Integer
    Dim br As Integer
    wr = sh.Range("C" & Application.Rows.Count).End(xlUp).Row
    br = sh.Range("D" & Application.Rows.Count).End(xlUp).Row
    whitePlayer = Sheet5.Range("C" & wr).Value
    blackPlayer = Sheet5.Range("D" & br).Value

    MsgBox whitePlayer
    MsgBox blackPlayer

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
