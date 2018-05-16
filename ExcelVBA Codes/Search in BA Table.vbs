Function CheckFileIsOpen(ByVal chkFile As String) As Boolean
    On Error Resume Next
    
    CheckFileIsOpen = (Workbooks(chkFile).Name = chkFile)
    
    On Error GoTo 0
    
End Function


Public Sub QryInBA(ByVal toFind As Variant)

    Dim destWS As Worksheet
    Dim destWB As Workbook
    Dim findResult As Range
    
    If False = CheckFileIsOpen("BA_2017_Hellriegel_KG.xls") Then
        Set destWB = Workbooks.Open("\\gh-dc\Global\Rechnungswesen\BA_2017_Hellriegel_KG.xls")
    Else
        Set destWB = Workbooks("BA_2017_Hellriegel_KG.xls")
    End If
    
    Set destWS = destWB.Worksheets("BA_neue_Firma")
        
    Set findResult = destWS.Cells.Find(What:=toFind, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False)
        
    If findResult Is Nothing Then
        MsgBox "Unable to find: " & toFind
    Else
        destWB.Activate
        destWS.Activate
        findResult.Activate
    End If

        
End Sub
