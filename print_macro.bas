Attribute VB_Name = "Module1"
Sub print_report()


    Dim ws As Worksheet
    Dim pdfPath As Variant

    Set ws = ThisWorkbook.Worksheets("Report")

    pdfPath = Application.GetSaveAsFilename( _
        InitialFileName:="Investment_Analysis_Report.pdf")

    If pdfPath = False Then Exit Sub

    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=CStr(pdfPath), _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=True, _
        OpenAfterPublish:=False

    MsgBox "Report PDF export worked."


End Sub

Sub Refresh_button()

    ActiveCell.Offset(-3, -2).Range("A1:A5").Select
    Selection.ClearContents

End Sub
