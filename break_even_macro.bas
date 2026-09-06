Attribute VB_Name = "Module3"
Sub Breakeven_analysis_macro()

    Dim ws As Worksheet
    
    Dim originalInvestment As Variant
    Dim originalPrice As Variant
    Dim originalVolume As Variant
    Dim originalInflation As Variant
    Dim originalDiscountRate As Variant
    
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    
    ' Store the current input values
    originalInvestment = ws.Range("Y5").Value
    originalPrice = ws.Range("Y12").Value
    originalVolume = ws.Range("Y11").Value
    originalInflation = ws.Range("Y9").Value
    originalDiscountRate = ws.Range("Y8").Value
    
    Application.ScreenUpdating = False
    
    On Error GoTo CleanUp
    
    ' --------------------------------------------------
    ' Break-even Initial Investment
    ' --------------------------------------------------
    ws.Range("H28").GoalSeek _
        Goal:=0, _
        ChangingCell:=ws.Range("Y5")
        
    ws.Range("Y28").Value = ws.Range("Y5").Value
    ws.Range("Y5").Value = originalInvestment
    
    
    ' --------------------------------------------------
    ' Break-even Sales Price
    ' --------------------------------------------------
    ws.Range("H28").GoalSeek _
        Goal:=0, _
        ChangingCell:=ws.Range("Y12")
        
    ws.Range("Y29").Value = ws.Range("Y12").Value
    ws.Range("Y12").Value = originalPrice
    
    
    ' --------------------------------------------------
    ' Break-even Sales Volume
    ' --------------------------------------------------
    ws.Range("H28").GoalSeek _
        Goal:=0, _
        ChangingCell:=ws.Range("Y11")
        
    ws.Range("Y30").Value = ws.Range("Y11").Value
    ws.Range("Y11").Value = originalVolume
    
    
    ' --------------------------------------------------
    ' Break-even Inflation
    ' --------------------------------------------------
    ws.Range("H28").GoalSeek _
        Goal:=0, _
        ChangingCell:=ws.Range("Y9")
        
    ws.Range("Y31").Value = ws.Range("Y9").Value
    ws.Range("Y9").Value = originalInflation
    
    
    ' --------------------------------------------------
    ' Break-even Discount Rate
    ' --------------------------------------------------
    ws.Range("H28").GoalSeek _
        Goal:=0, _
        ChangingCell:=ws.Range("Y8")
        
    ws.Range("Y32").Value = ws.Range("Y8").Value
    ws.Range("Y8").Value = originalDiscountRate
    
    
CleanUp:

    ' Always restore original inputs
    ws.Range("Y5").Value = originalInvestment
    ws.Range("Y12").Value = originalPrice
    ws.Range("Y11").Value = originalVolume
    ws.Range("Y9").Value = originalInflation
    ws.Range("Y8").Value = originalDiscountRate
    
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then
        MsgBox "Break-even analysis could not be completed." & vbCrLf & _
               "Error: " & Err.Description, _
               vbExclamation, "Break-even Analysis"
    Else
        MsgBox "Break-even analysis updated successfully.", _
               vbInformation, "Break-even Analysis"
    End If

End Sub

