Public Function fnQryRGInfo(ByVal strRGNr As String) As ADODB.Recordset

    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    conn.Open "Provider=sqloledb;Data Source=.\SQLEXPRESS;Initial Catalog=WWS_MIR;Integrated Security=SSPI"
    rs.Open "EXEC usp_RGInfo_by_Row " + strRGNr, conn, adOpenStatic
    
    Set fnQryRGInfo = rs
    
'Server=myServerAddress;Database=myDataBase;Trusted_Connection=True;
'Provider=sqloledb;Data Source=myServerName;Initial Catalog=myDatabaseName;Integrated Security=SSPI

End Function

Public Function fnFillRecordSet(ByRef rs As ADODB.Recordset)

    Dim curRow As Integer
    Dim curCol As Integer
    
    curRow = 3
    
    ' RecordCount is -1....
    With rs
'        If .RecordCount <> 0 Then
            
            'For curRow = 3 To (3 + .RecordCount)
                'For curCol = 1 To .Fields.Count
                    'ActiveSheet.Cells(curRow, curCol).Value = .Fields(curCol - 1).Value
                'Next curCol
'
                '.MoveNext
'
            'Next curRow
'
        'End If
        
        Do While Not rs.EOF
        
            For curCol = 1 To .Fields.Count
                ActiveSheet.Cells(curRow, curCol).Value = .Fields(curCol - 1).Value
            Next curCol
            
            
            curRow = curRow + 1
            rs.MoveNext
        Loop
    End With

End Function

Public Sub GetToC()

    If True = IsEmpty(ActiveCell.Value) Then
        MsgBox "unable qry without value"
        Exit Sub
    End If

    Dim rs As ADODB.Recordset
    
    Set rs = fnQryRGInfo(ActiveCell.Value)
    
    fnFillRecordSet rs

End Sub

Public Sub GetToB()
    
    If True = IsEmpty(ActiveCell.Value) Then
        MsgBox "unable qry without value"
        Exit Sub
    End If
    
    'Dim rs As ADODB.Recordset
    'Set rs = fnQryRGInfo(ActiveCell.Value)
    Dim strRGNr As String
    strRGNr = ActiveCell.Value
    
    
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    'conn.Open "Driver=sqlcli;Server=.\SQLEXPRESS;Database=WWS_MIR;Trusted_Connection=True;"
    conn.Open "Provider=sqloledb;Data Source=.\SQLEXPRESS;Initial Catalog=WWS_MIR;Integrated Security=SSPI"
    rs.Open "EXEC usp_RGInfo_by_Row " + strRGNr, conn, adOpenStatic
    
    'fnQryRGInfo = rs
    'ActiveSheet.Range("B1").CopyFromRecrodset rs
    
    Dim curRow As Integer
    Dim curCol As Integer
    
    curRow = 3
    
    ' RecordCount is -1....
    With rs
'        If .RecordCount <> 0 Then
            
            'For curRow = 3 To (3 + .RecordCount)
                'For curCol = 1 To .Fields.Count
                    'ActiveSheet.Cells(curRow, curCol).Value = .Fields(curCol - 1).Value
                'Next curCol
'
                '.MoveNext
'
            'Next curRow
'
        'End If
        
        Do While Not rs.EOF
        
            For curCol = 1 To .Fields.Count
                ActiveSheet.Cells(curRow, curCol).Value = .Fields(curCol - 1).Value
            Next curCol
            
            
            curRow = curRow + 1
            rs.MoveNext
        Loop
    End With
    

End Sub
