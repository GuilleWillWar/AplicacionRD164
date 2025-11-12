Attribute VB_Name = "FuncionesBotonesMacros"
Sub IrPrimeraHoja()
Attribute IrPrimeraHoja.VB_ProcData.VB_Invoke_Func = "p\n14"
    Sheets(1).Select
End Sub

Sub IrUltimaHoja()
Attribute IrUltimaHoja.VB_ProcData.VB_Invoke_Func = "u\n14"
    ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count).Activate
End Sub
