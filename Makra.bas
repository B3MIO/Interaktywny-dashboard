Attribute VB_Name = "Module1"
Public Sub OdswiezRaport()
    On Error GoTo ErrHandler

    ThisWorkbook.RefreshAll

    Dim ws As Worksheet, pt As PivotTable
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws

    Dim ch As ChartObject
    For Each ws In ThisWorkbook.Worksheets
        For Each ch In ws.ChartObjects
            ch.Chart.Refresh
        Next ch
    Next ws

    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Szablon zostal zaktualizowany.", vbInformation, "MsEX"
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Blad odswiezania szablonu: " & Err.Description, vbCritical, "MsEX"
End Sub



Public Sub EksportPDF()
    On Error GoTo ErrHandler
    Dim sciezka As String
    If ThisWorkbook.Path = "" Then
        MsgBox "Zapisz plik przed eksportem.", vbExclamation, "MsEX"
        Exit Sub
    End If

    sciezka = ThisWorkbook.Path & "\Raporty\Raport_" & Format(Now, "yyyy-mm-dd_hhmmss") & ".pdf"
    If Dir(ThisWorkbook.Path & "\Raporty", vbDirectory) = "" Then
        MkDir ThisWorkbook.Path & "\Raporty"
    End If

    Sheets("Raport").ExportAsFixedFormat Type:=xlTypePDF, Filename:=sciezka, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

    MsgBox "Zapisano raport jako PDF:" & vbCrLf & sciezka, vbInformation, "MsEX"
    Exit Sub

ErrHandler:
    MsgBox "Blad eksportu PDF: " & Err.Description, vbCritical, "MsEX"
End Sub

