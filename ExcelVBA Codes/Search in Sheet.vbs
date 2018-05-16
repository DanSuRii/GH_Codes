Sub Makro2()
'
' Makro2 Makro
'

'
    ' Selection.Copy
    Dim toSearch As Variant, findResult As Range, previousRng As Range, previousSheet As Worksheet
    
    Set previousRng = ActiveCell
    Set previousSheet = ActiveSheet
    
    toSearch = ActiveCell.Value2
    Sheets("BAListe").Select
    Range("A1").Select

    Set findResult = Cells.Find(What:=toSearch, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False)
    
    If findResult Is Nothing Then
        previousSheet.Activate
        previousRng.Activate
        MsgBox "Unable to find:" & toSearch
    Else
        findResult.Activate
    End If
End Sub