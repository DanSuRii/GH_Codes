Public Const G_WWS_Daten_Tbl_Name = "Tbl_WWS_Daten"
Public Const G_WWS_Daten_Sheet_Name = "Daten"

Public Function AddColumns(ByRef oWS As Worksheet)
    
    Dim colDaten As Variant
    
    colDaten = Array( _
        Array("ZEDatum", "=[@Netto]+10") _
        , Array("ZEKW", "= ""KW"" & TEXT( WEEKNUM([@ZEDatum],21), ""00"" ) & ""/"" & YEAR([@ZEDatum])") _
        , Array("Dauer ins Monat", "=MAX( MONTH([@ZEDatum])+ ( ( YEAR([@ZEDatum])-YEAR([@BelegDat]) ) * 12) - MONTH([@BelegDat]), 0 )") _
        , Array("BelegeDatum", "=DATEVALUE([@BelegDat])", "m/d/yyyy") _
        , Array("BelegJahrMon", "=CONCATENATE(YEAR([@BelegDat]),"" / "",TEXT( MONTH([@BelegDat]),""00""))") _
        , Array("Buch_Typ", "=IF([@[Betrag ()]]>0,""Soll"",""Haben"")") _
        , Array("ZEJahrMon", "=TEXT([@ZEDatum],""jjjj/MM"")") _
    )
        
        ' , Array("", "") _

    
    For Each curArr In colDaten
    
'        For Each curDat In curArr
        
            colName = curArr(0)
            colFormula = curArr(1)
        
            With oWS.ListObjects(G_WWS_Daten_Tbl_Name).ListColumns.Add
                .Name = colName
            End With
            
            Range(G_WWS_Daten_Tbl_Name & "[" & colName & "]").FormulaR1C1 = colFormula
            
        'Next curDat
    
    Next curArr

End Function

Public Sub testReplace()

    Dim varSubsetDaten As Variant
    
    varSubsetDaten = Array("Edeka", "Markant", "Rewe", "Dritte")

    sqlQry = "SELECT * FROM [WWS_MIR].[dbo].[vw_Liquidit?splanung_v2] WHERE [RG-NR] <= ( SELECT MAX(RLRENR) FROM WWS_MIR.dbo.Tbl_WWS_HRELEP WHERE RLREDA < 20180500 ) AND [KD-GRP] = N'?'  ORDER BY [RG-NR]"
    
    For Each cur In varSubsetDaten
    
        Debug.Print Replace(sqlQry, "?", cur)
    
    Next cur
    

End Sub

Public Function AddSubsets(ByRef oWB As Workbook)


    Dim varSubsetDaten As Variant
    
    varSubsetDaten = Array("Edeka", "Markant", "Rewe", "Dritte")
    
    sqlQry = "SELECT * FROM [WWS_MIR].[dbo].[vw_Liquidit?splanung_v2] WHERE [RG-NR] <= ( SELECT MAX(RLRENR) FROM WWS_MIR.dbo.Tbl_WWS_HRELEP WHERE RLREDA < 20180500 ) AND [KD-GRP] = N'?'  ORDER BY [RG-NR]"
    
    For Each cur In varSubsetDaten
    
    addQryResult oWB _
    , Replace(sqlQry, "?", cur) _
    , cur _
    , "Tbl_WWS_" & cur
    
    Next cur



End Function

Public Function addQryResult(ByRef oWB As Workbook, ByVal sqlQry As String, ByVal sheetName As String, ByVal tblName As String)

    connstring = "ODBC;DRIVER=SQL Server Native Client 11.0;SERVER=.\SQLEXPRESS;UID=A.Roennburg;Trusted_Connection=Yes;APP=Microsoft Office 2013;WSID="
    connproperty = "LOGISTIK-8;ServerSPN=WWS_MIR;"
    
    Dim oWS As Worksheet
    Set oWS = oWB.Worksheets.Add
    oWS.Name = sheetName
    
    With oWS.ListObjects.Add(SourceType:=0, Source:=Array(Array( _
        connstring _
        ), Array(connproperty)), Destination:=Range("A1")). _
        QueryTable
        
        .CommandText = Array( _
         sqlQry _
        )
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = tblName
        .ListObject.TableStyle = "TableStyleLight1"
        .Refresh BackgroundQuery:=False
    
    End With
    
    
End Function


Public Sub doGenerate()

    Dim oWB As Workbook
    Dim oWS As Worksheet
    Dim rngDBData As Range
    
    Set oWB = Workbooks.Add
    Set oWS = oWB.Sheets(1)
    
    oWS.Name = G_WWS_Daten_Sheet_Name
    
    Set rngDBData = oWS.Range("A1")
    
    ' Connection ...... SQL Query.... Display Name
    
    sqlQry = "SELECT * FROM [WWS_MIR].[dbo].[vw_Liquidit?splanung_v2] WHERE [RG-NR] <= ( SELECT MAX(RLRENR) FROM WWS_MIR.dbo.Tbl_WWS_HRELEP WHERE RLREDA < 20180500 ) ORDER BY [RG-NR]"
    connstring = "ODBC;DRIVER=SQL Server Native Client 11.0;SERVER=.\SQLEXPRESS;UID=A.Roennburg;Trusted_Connection=Yes;APP=Microsoft Office 2013;WSID="
    
    With oWS.ListObjects.Add(SourceType:=0, Source:=Array(Array( _
        "ODBC;DRIVER=SQL Server Native Client 11.0;SERVER=.\SQLEXPRESS;UID=A.Roennburg;Trusted_Connection=Yes;APP=Microsoft Office 2013;WSID=" _
        ), Array("LOGISTIK-8;ServerSPN=WWS_MIR;")), Destination:=Range("A1")). _
        QueryTable
    
        .CommandText = Array( _
            sqlQry _
        )
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "Tbl_WWS_Daten"
        .ListObject.TableStyle = "TableStyleLight1"
        .Refresh BackgroundQuery:=False
    End With
    
    AddColumns oWS
    AddSubsets oWB
    
    
'    With oWS.QueryTables.Add(Connection:=connstring, Destination:=rngDBData, Sql:=sqlQry)
'        .Name = "Qry_WWS_Liquid"
'        .Refresh
'
'    End With


    ' oWB.Close

End Sub


Public Sub forArrTest2()

    Dim vatest As Variant
    
    vatest = Array( _
        Array("ZEDatum", "=[@Netto]+10") _
        , Array("ZEKW", "= ""KW"" & TEXT( KALENDERWOCHE([@ZEDatum];21); ""00"" ) & ""/"" & JAHR([@ZEDatum])") _
        , Array("Dauer ins Monat", "=MAX( MONAT([@ZEDatum])+ ( ( JAHR([@ZEDatum])-JAHR([@BelegDat]) ) * 12) - MONAT([@BelegDat]); 0 )") _
        , Array("BelegJahrMon", "=WENN([@[Betrag ()]]>0;""Soll"";""Haben"")") _
        , Array("Buch_Typ", "=WENN([@[Betrag ()]]>0;""Soll"";""Haben"")") _
        , Array("ZEJahrMon", "=TEXT([@ZEDatum];""jjjj/MM"")") _
        )

    

    
    For Each curArr In vatest
    
        For Each curDat In curArr
            Debug.Print curDat
        Next curDat
    
    Next curArr


End Sub

Public Sub forArrTest()

    Dim colDaten As Variant
    
    colDaten = Array( _
        Array("ZEDatum", "=[@Netto]+10") _
        , Array("ZEKW", "= ""KW"" & TEXT( KALENDERWOCHE([@ZEDatum];21); ""00"" ) & ""/"" & JAHR([@ZEDatum])") _
        , Array("Dauer ins Monat", "=MAX( MONAT([@ZEDatum])+ ( ( JAHR([@ZEDatum])-JAHR([@BelegDat]) ) * 12) - MONAT([@BelegDat]); 0 )") _
        , Array("Belegedatum", "=DATWERT([@BelegDat])") _
        , Array("BelegJahrMon", "=VERKETTEN(JAHR([@BelegDat]);"" / "";TEXT( MONAT([@BelegDat]);""00""))") _
        , Array("Buch_Typ", "=WENN([@[Betrag ()]]>0;""Soll"";""Haben"")") _
        , Array("ZEJahrMon", "=TEXT([@ZEDatum];""jjjj/MM"")") _
    )
        ' , Array("", "") _

    For Each curArr In colDaten
    
        For Each curDat In curArr
            Debug.Print curDat
        Next curDat
    
    Next curArr


End Sub

