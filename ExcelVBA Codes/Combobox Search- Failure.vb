
'Private Sub ComboBox1_GotFocus()
 '      Dim SearchKey As Object
  '     Set SearchKey = ActiveWorkbook.Names("SearchKey")
       
       
         ' ComboBox1.Clear
        ' Dim SearchRange As Range
        ' Set Range = Range("Artikel!SearchKey")
        'ComboBox1.ListFillRange = "Artikel!F:F"
        ' Me.ComboBox1.DropDown
'End Sub

Private Sub ComboBox1_Change()
    
    If ComboBox1.Value = "" Then
        Exit Sub
    End If
    
    ' Application.ScreenUpdating = False
    
    ' SendKeys "{ESC}"
    
    
    Dim strSQL As String
    strSQL = "SELECT SearchKey From [Artikel$] " & _
            " WHERE SearchKey Like '%" & ComboBox1.Value & "%'"
            
    Dim strConn As String
    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""" & ThisWorkbook.Path & "\" & _
            ActiveWorkbook.Name & """;" & "Extended Properties=Excel 12.0 Macro;"

    Dim rs As New ADODB.Recordset
    
    rs.Open strSQL, strConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
'     ComboBox1.Clear
        
    If Not rs.EOF Then
    
        With ComboBox1
            .Clear
        Do
            Me.ComboBox1.AddItem rs![SearchKey]
            rs.MoveNext
        Loop Until rs.EOF
        End With
        
        
        'While Not rs.EOF
            'rs.MoveNext
            ' debugprint rs.GetString
        'Wend
    
    End If
    
    rs.Close
    Set rs = Nothing
    
    ' Application.ScreenUpdating = True
    
End Sub


