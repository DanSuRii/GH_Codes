Sub SetPrintAreaToTable()
'
' Makro4 Makro
'

'
'    Dim prinarea As String
    Dim TableA As ListObject
    Set TableA = ActiveSheet.ListObjects("Tabelle1")
'    PrintArea = Range("Tabelle1[#Alle]").Address

    PrintArea = TableA.Range.Address

'    Range("Tabelle1[#All]").Select
'    Range("I14").Activate
    ActiveSheet.PageSetup.PrintArea = PrintArea
End Sub


Sub Makro1()

    Dim FileName As String
    Dim fso
    
    Set fso = CreateObject("Scripting.FileSystemObject")
'    FileName = ActiveWorkbook.Name
'    Dim fso As New scripting.filesystemobject
    FileName = fso.GetBaseName(ActiveWorkbook.Name)
    
    SetPrintAreaToTable

'    ActiveSheet.PageSetup.PrintArea = "Tabelle1[#Alle]"
    
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.787401575)
        .BottomMargin = Application.InchesToPoints(0.787401575)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 0
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
        "C:\Users\A.Roennburg\Documents\GH_ArbeitPlatz\Arbeits vom Sandra\MARKANT_CalculateSheet\PDF Daten\" & FileName & ".pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=True
'    ChDir _
'        "C:\Users\A.Roennburg\Documents\GH_ArbeitPlatz\Arbeits vom Sandra\MARKANT_CalculateSheet\CSV Daten"
        
    ExportCSV FileName
    
    
End Sub



Public Sub ExportCSV(ByVal FileName As String)
    ' If FileName Is Null Then FileName = ActiveSheet.Name
    
    Range("Tabelle1[#All]").Select
'    Range("D13").Activate
    Selection.Copy
    Workbooks.Add
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs FileName:= _
        "C:\Users\A.Roennburg\Documents\GH_ArbeitPlatz\Arbeits vom Sandra\MARKANT_CalculateSheet\CSV Daten\J10003.csv" _
        , FileFormat:=xlCSV, CreateBackup:=False, Local:=True
'    ActiveWorkbook.Save
    ActiveWindow.Close False
    
    Application.DisplayAlerts = True

End Sub
