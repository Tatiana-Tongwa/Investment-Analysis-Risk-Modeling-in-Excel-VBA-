Attribute VB_Name = "Module4"
Sub macro_monte_carlo()
Attribute macro_monte_carlo.VB_ProcData.VB_Invoke_Func = " \n14"
'
' macro_monte_carlo Macro
'

'
Sheets("MC_hist").EnableCalculation = False
Sheets("MC_data").Range("A:A").ClearContents
n = Sheets("MC_sim").Range("C12")

For i = 1 To n

Sheets("MC_data").Range("A" & i) = Sheets("Sheet1").Range("H71")
Sheets("MC_data").Range("B" & i) = Sheets("Sheet1").Range("H72")
Sheets("MC_data").Range("C" & i) = Sheets("Sheet1").Range("H73")
Sheets("MC_data").Range("D" & i) = Sheets("Sheet1").Range("H74")
Sheets("MC_data").Range("E" & i) = Sheets("Sheet1").Range("H75")

Next
Sheets("MC_hist").EnableCalculation = True

End Sub
