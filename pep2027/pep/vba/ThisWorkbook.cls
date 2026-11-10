Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    Application.EnableEvents = True
    ThisWorkbook.Names("Makrostatus").RefersToRange.Calculate
End Sub
