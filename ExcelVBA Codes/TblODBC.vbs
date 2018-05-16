

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

Public Sub AddNewQryTbl()
    sqlQry = "SELECT * FROM [WWS_MIR].[dbo].[vw_Liquiditätsplanung_v2] ORDER BY [RG-NR]"
    connstring = "ODBC;DRIVER=SQL Server Native Client 11.0;SERVER=.\SQLEXPRESS;UID=A.Roennburg;Trusted_Connection=Yes;APP=Microsoft Office 2013;WSID="
  With ActiveSheet.QueryTables.Add(Connection:=connstring, Destination:=Range("A1"), Sql:=sqlQry)
    .Refresh
  End With

End Sub

Sub Makro3()
'
' Makro3 Makro
'
    Range("Tabelle_Abfrage_von_WWSMIR[Betrag (€)]").Select
    Selection.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "

'    ActiveSheet.QueryTables(1)

End Sub

Type Dahn_Column
    Name As String
    Formula As String
End Type

Public Function AddToCollection( ByRef Cont as Collection, ByVal Name as string , ByVal Formula as string)
    dim newCol as new Dahn_Column
    with newCol 
        .Name = Name
        .Formula = Formula
    end with

    Cont.Add newCol

End Function

Public Sub AddColumn()
    Dim ColumnArr as variant
    Dim FormulaArr as variant
    
    ColumnArr = Array("ZEDatum",	"ZEKW", "Dauer ins Monat",	"Belegedatum",	"BelegJahrMon",	"Buch_Typ",	"ZEJahrMon");
    FormulaArr = Array(
    "=[@Netto]+10"
    , "= ""KW"" & TEXT( KALENDERWOCHE([@ZEDatum];21); ""00"" ) & ""/"" & JAHR([@ZEDatum])" _
    ,"=MAX( MONAT([@ZEDatum])+ ( ( JAHR([@ZEDatum])-JAHR([@BelegDat]) ) * 12) - MONAT([@BelegDat]); 0 )" _
    ,"=DATWERT([@BelegDat])	=VERKETTEN(JAHR([@BelegDat]);"/";TEXT( MONAT([@BelegDat]);""00""))" _
    ,"=WENN([@[Betrag (€)]]>0;""Soll"";""Haben"")" _
    ,"=TEXT([@ZEDatum];""jjjj/MM"")"
    );


    dim dColCont as new Collection
    AddToCollection dColCont,  "ZEDatum", "=[@Netto]+10" 
    AddToCollection dColCont, "ZEKW", "= ""KW"" & TEXT( KALENDERWOCHE([@ZEDatum];21); ""00"" ) & ""/"" & JAHR([@ZEDatum])" 
    AddToCollection dColCont, "Dauer ins Monat", "=MAX( MONAT([@ZEDatum])+ ( ( JAHR([@ZEDatum])-JAHR([@BelegDat]) ) * 12) - MONAT([@BelegDat]); 0 )" 
    AddToCollection dColCont, "BelegJahrMon", "=DATWERT([@BelegDat])	=VERKETTEN(JAHR([@BelegDat]);"/";TEXT( MONAT([@BelegDat]);""00""))" 
    AddToCollection dColCont, "Buch_Typ", "=WENN([@[Betrag (€)]]>0;""Soll"";""Haben"")" 
    AddToCollection dColCont, "ZEJahrMon", "=TEXT([@ZEDatum];""jjjj/MM"")" 
    




    for each curCol in dColCont
        With ActiveSheet.ListObjects("Tabelle_Abfrage_von_WWSMIR").ListColumns.Add
            .Name = curCol.Name
        End With
        
        Range("Tabelle_Abfrage_von_WWSMIR[" & curCol.Name  & "]").FormulaR1C1 = _
            curCol.Formula

    next curCol

    Range("Tabelle_Abfrage_von_WWSMIR[ZEDatum]").FormulaR1C1 = _
        "=Tabelle_Abfrage_von_WWSMIR[@Netto]+10"
    Range("Tabelle_Abfrage_von_WWSMIR[ZEDatum]").Select
        Selection.NumberFormat = "m/d/yyyy"

End Sub

