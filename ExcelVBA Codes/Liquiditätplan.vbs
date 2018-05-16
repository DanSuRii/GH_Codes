Sub Makro1()
'
' Makro1 Makro
'

'
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:=Array(Array( _
        "ODBC;DRIVER=SQL Server Native Client 11.0;SERVER=.\SQLEXPRESS;UID=A.Roennburg;Trusted_Connection=Yes;APP=Microsoft Office 2013;WSID=" _
        ), Array("LOGISTIK-8;ServerSPN=WWS_MIR;")), Destination:=Range("$A$1")). _
        QueryTable
       '  .CommandType = 0
        .CommandText = Array( _
        "SELECT *" & Chr(13) & "" & Chr(10) & "FROM [WWS_MIR].[dbo].[vw_Liquiditätsplanung_v2]" & Chr(13) & "" & Chr(10) & "  ORDER BY [RG-NR]" _
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
        .ListObject.DisplayName = "Tabelle_Abfrage_von_WWSMIR"
        .Refresh BackgroundQuery:=False
    End With
End Sub

Public Sub AddNewQryTbl()
    sqlQry = "SELECT * FROM [WWS_MIR].[dbo].[vw_Liquiditätsplanung_v2] ORDER BY [RG-NR]"
    connstring = "ODBC;DRIVER=SQL Server Native Client 11.0;SERVER=.\SQLEXPRESS;UID=A.Roennburg;Trusted_Connection=Yes;APP=Microsoft Office 2013;WSID="
  With ActiveSheet.QueryTables.Add(Connection:=connstring, Destination:=Range("A1"), Sql:=sqlQry)
    .Refresh
  End With

End Sub


Sub Makro2()
'
' Makro2 Makro
'

'
    With ActiveWorkbook.Connections("Abfrage von WWSMIR").ODBCConnection
        .BackgroundQuery = True
        .CommandText = Array( _
        "SELECT *" & Chr(13) & "" & Chr(10) & "  FROM [WWS_MIR].[dbo].[vw_RG_KDST_JIT_einf]")
        .CommandType = xlCmdSql
        .Connection = Array(Array( _
        "ODBC;DRIVER=SQL Server Native Client 11.0;SERVER=.\SQLEXPRESS;UID=A.Roennburg;Trusted_Connection=Yes;APP=Microsoft Office 2013;WSID=" _
        ), Array("LOGISTIK-8;ServerSPN=WWS_MIR;"))
        .RefreshOnFileOpen = False
        .SavePassword = False
        .SourceConnectionFile = ""
        .SourceDataFile = ""
        .ServerCredentialsMethod = xlCredentialsMethodIntegrated
        .AlwaysUseConnectionFile = False
    End With
    With ActiveWorkbook.Connections("Abfrage von WWSMIR")
        .Name = "Abfrage von WWSMIR"
        .Description = ""
    End With
    ActiveWorkbook.Connections("Abfrage von WWSMIR").Refresh
    With Selection.ListObject.QueryTable
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
    End With
End Sub



Public Sub GetQryTbl()
    QryTbl = ActiveSheet.QueryTables(1)
    
End Sub


Public Sub AddColumn()
    With ActiveSheet.ListObjects("Tabelle_Abfrage_von_WWSMIR").ListColumns.Add
        .Name = "Belegedatum"
    End With
End Sub

Sub Makro3()
'
' Makro3 Makro
'

'
    Range("Tabelle_Abfrage_von_WWSMIR[Betrag (€)]").Select
    Selection.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
End Sub
Sub Makro4()
'
' Makro4 Makro
'

'
    ActiveCell.FormulaR1C1 = "ZEDatum"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=[@Netto]+10"
    Range("I3").Select
    Range("Tabelle_Abfrage_von_WWSMIR[ZEDatum]").FormulaR1C1 = _
        "=Tabelle_Abfrage_von_WWSMIR[@Netto]+10"
    Range("Tabelle_Abfrage_von_WWSMIR[ZEDatum]").Select
    Selection.NumberFormat = "m/d/yyyy"
End Sub

Sub Makro5()
'
' Makro5 Makro
'

'
    Range("Tabelle_Abfrage_von_WWSMIR[[#Headers],[ZEDatum]]").Select
    Selection.ListObject.ListColumns.Add
    Range("Tabelle_Abfrage_von_WWSMIR[[#Headers],[Spalte1]]").Select
'    Windows("KW Plannung 05042018.xlsx").Activate
'    Windows("Mappe6").Activate
    ActiveCell.FormulaR1C1 = "ZEKW"
    Range("J2").Select
End Sub


Public Function AddToCollection(ByRef Cont As Collection, ByVal Name As String, ByVal Formula As String, Optional Format As String = "")
    Dim newCol As New Dahn_Col
    With newCol
        .Name = Name
        .Formula = Formula
        .Format = Format
    End With

    Cont.Add newCol

End Function

Public Sub AddColumn()


    Dim dColCont As New Collection
    AddToCollection Cont:=dColCont, Name:="ZEDatum", Formula:="=[@Netto]+10", Format:="m/d/yyyy"
    AddToCollection dColCont, "ZEKW", "= ""KW"" & TEXT( WEEKNUM([@ZEDatum],21), ""00"" ) & ""/"" & YEAR([@ZEDatum])"
    AddToCollection dColCont, "Dauer ins Monat", "=MAX( MONTH([@ZEDatum])+ ( ( YEAR([@ZEDatum])-YEAR([@BelegDat]) ) * 12) - MONTH([@BelegDat]), 0 )"
    AddToCollection dColCont, "BelegeDatum", "=DATEVALUE([@BelegDat])", "m/d/yyyy"
    AddToCollection dColCont, "BelegJahrMon", "=CONCATENATE(YEAR([@BelegDat]),"" / "",TEXT( MONTH([@BelegDat]),""00""))"
    AddToCollection dColCont, "Buch_Typ", "=IF([@[Betrag (€)]]>0,""Soll"",""Haben"")"
    AddToCollection dColCont, "ZEJahrMon", "=TEXT([@ZEDatum],""jjjj/MM"")"
    


    For Each curCol In dColCont
        With ActiveSheet.ListObjects("Tabelle_Abfrage_von_WWSMIR").ListColumns.Add
            .Name = curCol.Name
        End With
        
        tmpFormula = curCol.Formula
        
        With Range("Tabelle_Abfrage_von_WWSMIR[" & curCol.Name & "]")
            .FormulaR1C1 = curCol.Formula
            .NumberFormat = curCol.Format
        End With

    Next curCol


End Sub




Public Sub CopyAt()

    Dim oWB As Workbook, oDestWB As Workbook
    Dim oWS As Worksheet, oDestWS As Worksheet
    
    Dim rngSrc As Range, rngDest As Range
    

    Set oWB = Workbooks("Generate automatic.xlsm")
    Set oWS = oWB.Worksheets("Tabelle2")
    Set rngSrc = oWS.UsedRange
    
    Set oDestWB = Workbooks("KW Plannung 03052018.xlsx")
    Set oDestWS = oDestWB.Worksheets("Daten")
    Set rngDest = oDestWS.Range("A3")
    
    rngSrc.Copy rngDest
    
    Dim rngUsed As Range
    
    Set rngUsed = oDestWS.UsedRange
    
    oDestWB.Activate
    oDestWS.Activate
    
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    
    Selection.Cut Destination:=Range("A2")
    
    'Range("A4:H12556").Cut Destination:=Range("A2:H12554")
    'Range("Tabelle_Abfrage_von_WWSMIR[[KD-WWS]:[RG_DESC]]").Select
    'Windows("Generate automatic.xlsm").Activate
    

End Sub


Sub Makro1()
'
' Makro1 Makro
'

'
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$H$12478"), , xlYes).Name _
        = "Tabelle_ExterneDaten_1"
    Range("Tabelle_ExterneDaten_1[#All]").Select
    Windows("KW Plannung 03052018.xlsx").Activate
    Windows("Mappe6").Activate
End Sub

Sub makro33()
    ActiveSheet.ListObjects("Tbl_WWS_Daten").TableStyle =   
end Sub