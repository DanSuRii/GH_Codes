

Public Sub InsertValue()
    Dim oWS As Worksheet
    Dim oTbl As ListObject
    Dim newRow As ListRow
    Dim Orgn As String
    Dim MwST As Double, betr As Double
    Dim WKD_MKT As Double, WKD_KD As Double, DL As Double
    
    Set oWS = ActiveWorkbook.Worksheets("ARBETISTABELLE")
    Set oTbl = oWS.ListObjects("Tabelle1")
    Set newRow = oTbl.ListRows.Add
    
    Orgn = cbOrgn.value
    
    betr = RetIfValue(tbBetr.value)
    MwST = Format(Me.cbPCT.value, "0.00")
    DL = RetIfValue(Me.tbDL.value)
    WKD_MKT = RetIfValue(Me.tbWKDMK.value)
    WKD_KD = RetIfValue(Me.tbWKDKD.value)
    
    With newRow.Range
        .Cells(1, oTbl.ListColumns("KDORGN_NAM").Index) = Orgn
        .Cells(1, oTbl.ListColumns("RGNr").Index) = tbRGNR.value
        .Cells(1, oTbl.ListColumns("RGBETR").Index) = betr
        .Cells(1, oTbl.ListColumns("MwSTPCT").Index) = MwST
        .Cells(1, oTbl.ListColumns("WKD_Markant").Index) = WKD_MKT
        If WKD_KD <> 0 Then .Cells(1, oTbl.ListColumns("WKD_Kund(HSI)").Index) = WKD_KD
        .Cells(1, oTbl.ListColumns("DL").Index) = DL
    End With
    
End Sub

Public Function RetIfValue(ByVal value) As Double
    Dim ret As Double
    ret = 0
    
    If IsNumeric(value) Then ret = CDbl(value)
    
    RetIfValue = ret

End Function

Private Sub btnEinfg_Click()
    InsertValue
End Sub


Private Sub cbPCT_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    cbPCT.value = Format(cbPCT.value, "0%")

End Sub
