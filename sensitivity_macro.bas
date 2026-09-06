Attribute VB_Name = "Module11"
Sub Sensitivity()

    Dim ws As Worksheet
    Dim rowNum As Long
    Dim changeStep As Long
    Dim sensitivityFactor As Double
    
    Set ws = ThisWorkbook.Worksheets("sensitivity")
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' ---------------------------------------------------------
    ' Run sensitivity analysis for the five variables
    ' Rows 3 to 7
    ' ---------------------------------------------------------
    
    For rowNum = 3 To 7
        
        ' Test changes from +50% to -50%
        For changeStep = 5 To -5 Step -1
            
            sensitivityFactor = 1 + (changeStep / 10#)
            
            ' Update sensitivity coefficient
            ws.Cells(rowNum, 2).Value = sensitivityFactor
            
            ' Recalculate model
            Application.Calculate
            
            ' Do not overwrite the base-case formula columns:
            ' H = base NPV
            ' T = base PI
            If changeStep <> 0 Then
                
                ' NPV sensitivity
                ' H is column 8
                ws.Cells(rowNum, 8 + changeStep).Value = _
                    ws.Cells(rowNum, 8).Value
                
                ' Profitability Index sensitivity
                ' T is column 20
                ws.Cells(rowNum, 20 + changeStep).Value = _
                    ws.Cells(rowNum, 20).Value
                
            End If
            
        Next changeStep
        
        ' Return sensitivity coefficient to base case
        ws.Cells(rowNum, 2).Value = 1
        
    Next rowNum
    
    ' Recalculate original base case
    Application.Calculate
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "Sensitivity analysis updated successfully.", _
           vbInformation, "Sensitivity Analysis"
    
    Exit Sub


ErrorHandler:

    ' Restore sensitivity coefficients
    For rowNum = 3 To 7
        ws.Cells(rowNum, 2).Value = 1
    Next rowNum
    
    Application.Calculate
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "Sensitivity analysis could not be completed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbExclamation, "Sensitivity Analysis"

End Sub

