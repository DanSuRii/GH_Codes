
Public Sub ClearTBs()
    Dim cCont As Control
    
    For Each cCont In Me.Controls
        If "TextBox" = TypeName(cCont) Then
            ' No one knows that is TextBox but Possible to Call...
            cCont.Text = ""
        End If
    Next cCont
End Sub

Private Sub btn_Clear_Click()
    Me.ClearTBs
    
End Sub

Public Sub InsertValue()
    Dim oWS As Worksheet
    Dim oLO As ListObject
    Dim myNewRow As ListRow
    Dim gesamt As Double, ZBetr As Double, neunzehn As Double
    
    
    
    
    Set Worksheet = ActiveWorkbook.Worksheets("TblRG")
    Set oLO = Worksheet.ListObjects("Tbl_RGListe")
    Set myNewRow = oLO.ListRows.Add
    
    If IsNumeric(Me.Controls("tb_Gesamt").Value) Then gesamt = CDbl(Me.Controls("tb_Gesamt").Value)
    If IsNumeric(Me.Controls("tb_Zahlbetr").Value) Then ZBetr = CDbl(Me.Controls("tb_Zahlbetr").Value)
    If IsNumeric(Me.Controls("tb_neunzehn").Value) Then neunzehn = CDbl(Me.Controls("tb_neunzehn").Value)
    
    
    myNewRow.Range.Cells(1, oLO.ListColumns("RGNr").Index) = Me.Controls("tb_RGNR").Value
    myNewRow.Range.Cells(1, oLO.ListColumns("Gesamt").Index) = gesamt
    myNewRow.Range.Cells(1, oLO.ListColumns("ZahlungBETR").Index) = ZBetr
    If neunzehn <> 0 Then myNewRow.Range.Cells(1, oLO.ListColumns("19%").Index) = neunzehn
    
    
End Sub

Private Sub btn_Einfg_Click()
    Me.InsertValue
End Sub

Private Sub tb_Gesamt_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Me.FnKeyWork KeyAscii
End Sub

Private Sub tb_neunzehn_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Me.FnKeyWork KeyAscii
End Sub

Private Sub tb_RGNR_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Me.FnKeyWork KeyAscii
End Sub

Private Sub tb_Zahlbetr_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Me.FnKeyWork KeyAscii
End Sub

Public Sub FnKeyWork(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = vbKeyF2 Then Me.InsertValue
    
End Sub
