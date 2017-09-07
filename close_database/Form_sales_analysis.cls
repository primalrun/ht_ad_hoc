VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sales_analysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub cmd_check_for_doc_numbers_Click()

Call GL.CheckForDocNumbers

End Sub



Private Sub cmd_check_status_Click()

Dim strsql As String
Dim qdefDAO As DAO.QueryDef

Set qdefDAO = CurrentDb.QueryDefs("status")
qdefDAO.SQL = "select * from global"
Set qdefDAO = Nothing

DoCmd.OpenQuery "status"

End Sub


Private Sub cmd_comparisons_Click()

Call Compare.Comparison

End Sub

Private Sub cmd_ecommerce_accrual_Click()

Dim appXL As Excel.Application
Dim wbCur As Excel.Workbook
Dim wsLogic As Excel.Worksheet

Call PublicStuff.GetFile("Select EC Accrual File")

If blnSelectFile = False Then
    MsgBox "No file selected, process cancelled", vbCritical + vbSystemModal, "File Validation"
    Exit Sub
End If

Call PublicStuff.SelectSaveDirectory("Select Directory to Save E-Commerce Accrual File")
If strPathSave = "" Then
    MsgBox "No Save Directory selected, process cancelled", vbCritical + vbSystemModal, "Save Directory Validation"
    Exit Sub
End If

Set appXL = New Excel.Application
appXL.Workbooks.Open (varFiles(1))
appXL.ActiveWorkbook.SaveAs strPathSave & "E-Commerce Accrual.xlsx", 51
appXL.ActiveSheet.Name = "Data"
Set wbCur = appXL.ActiveWorkbook
Set wsLogic = wbCur.Sheets.Add
wsLogic.Name = "Logic"
appXL.Cells(1, 1) = "Include all records where Qty_ Shipped Not Invoiced <> 0"
appXL.Cells(2, 1) = "The sum of Amount is used for Sales and AR"
appXL.Cells(3, 1) = "The sum of TotCost is used for COGS and Accrued Expenses"

appXL.ActiveWorkbook.Close True
appXL.Quit
Set appXL = Nothing

Call HT.ECommerceAccrual

MsgBox "Success: E-Commerce Accrual Data is Complete", vbInformation + vbSystemModal, "E-Commerce Accrual Data"

End Sub

Private Sub cmd_ht_dw_sales_analysis_Click()

Call HT.SalesAnalysisHT

End Sub

Private Sub cmd_internal_sales_ht_Click()

Call HT.InternalSalesHitouch

End Sub

Private Sub cmd_internal_sales_myop_Click()

Call MYOP.InternalSalesMYOP


End Sub

Private Sub cmd_line_variances_Click()

Call LineVariances.HTBuild

MsgBox "Success: Line Variance is Updated", vbInformation + vbSystemModal, "Line Variance"

End Sub

Private Sub cmd_montefiore_copier_Click()

DoCmd.OpenQuery "montefiore_copier", acViewNormal

End Sub

Private Sub cmd_myop_dw_sales_analysis_Click()

Call MYOP.SalesAnalysisMYOP

End Sub

Private Sub cmd_reclass_po_and_toner_Click()

Call PublicStuff.GetFile("Select PO and Toner File")

If blnSelectFile = False Then
    MsgBox "No file selected, process cancelled", vbCritical + vbSystemModal, "File Validation"
    Exit Sub
End If

Call HT.ReclassPOandToner

End Sub

Private Sub cmd_reclass_tsc_kitting_cogs_Click()

Call MYOP.ReclassTSCKittingCOGS

End Sub

Private Sub cmd_sales_and_cost_total_by_gl_company_Click()

DoCmd.OpenQuery "sales_cost_total_by_gl_company", acViewNormal

End Sub

Private Sub cmd_update_close_period_Click()

Dim dteStart As Date
Dim dteEnd As Date
Dim strYear As String
Dim strMonth As String

strYear = InputBox("Close Year? yyyy", "Close Year", Year(Date))

If IsNumeric(strYear) = False Or Len(strYear) <> 4 Then
    MsgBox "Year is not a 4 digit number, Process Cancelled", vbCritical + vbSystemModal, "Year Validation"
    Exit Sub
End If

strMonth = InputBox("Close Month Number? n", "Close Month Number", Month(Date) - 1)

Select Case True
    Case IsNumeric(strMonth) = False
        MsgBox "Month is not a number, Process Cancelled", vbCritical + vbSystemModal, "Month Validation"
        Exit Sub
    Case strMonth < 1
        MsgBox "Month is < 1, Process Cancelled", vbCritical + vbSystemModal, "Month Validation"
        Exit Sub
    Case strMonth > 12
        MsgBox "Month is > 1, Process Cancelled", vbCritical + vbSystemModal, "Month Validation"
        Exit Sub
End Select

dteStart = DateSerial(CInt(strYear), CInt(strMonth), 1)
dteEnd = DateAdd("d", -1, DateAdd("m", 1, dteStart))

CurrentDb.Execute "update global set value_date = #" & dteStart & "# where category = 'close_date_start'"
CurrentDb.Execute "update global set value_date = #" & dteEnd & "# where category = 'close_date_end'"

MsgBox "Complete", vbInformation + vbSystemModal, "Update Close Period"

End Sub

Private Sub cmd_update_dw_sales_product_data_Click()

Call DW.DWSalesProduct

End Sub

Private Sub cmd_update_gl_sales_cogs_Click()

Call GL.UpdateGLSalesandCOGS

End Sub

Private Sub cmd_check_for_doc_numbers_GotFocus()

Me.cmd_check_for_doc_numbers.BackColor = RGB(0, 176, 80)

End Sub

Private Sub cmd_check_for_doc_numbers_LostFocus()

Me.cmd_check_for_doc_numbers.BackColor = RGB(166, 166, 166)

End Sub

Private Sub cmd_check_status_GotFocus()

Me.cmd_check_status.BackColor = RGB(0, 176, 80)

End Sub

Private Sub cmd_check_status_LostFocus()

Me.cmd_check_status.BackColor = RGB(166, 166, 166)

End Sub

Private Sub cmd_update_close_period_GotFocus()

Me.cmd_update_close_period.BackColor = RGB(0, 176, 80)

End Sub

Private Sub cmd_update_close_period_LostFocus()

Me.cmd_update_close_period.BackColor = RGB(166, 166, 166)

End Sub

Private Sub cmd_update_dw_sales_product_data_GotFocus()

Me.cmd_update_dw_sales_product_data.BackColor = RGB(0, 176, 80)

End Sub

Private Sub cmd_update_dw_sales_product_data_LostFocus()

Me.cmd_update_dw_sales_product_data.BackColor = RGB(166, 166, 166)

End Sub

Private Sub cmd_ht_dw_sales_analysis_GotFocus()

Me.cmd_ht_dw_sales_analysis.BackColor = RGB(0, 176, 80)

End Sub

Private Sub cmd_ht_dw_sales_analysis_LostFocus()

Me.cmd_ht_dw_sales_analysis.BackColor = RGB(166, 166, 166)

End Sub

Private Sub cmd_myop_dw_sales_analysis_GotFocus()

Me.cmd_myop_dw_sales_analysis.BackColor = RGB(0, 176, 80)

End Sub

Private Sub cmd_myop_dw_sales_analysis_LostFocus()

Me.cmd_myop_dw_sales_analysis.BackColor = RGB(166, 166, 166)

End Sub

Private Sub cmd_ecommerce_accrual_GotFocus()

Me.cmd_ecommerce_accrual.BackColor = RGB(0, 176, 80)

End Sub

Private Sub cmd_ecommerce_accrual_LostFocus()

Me.cmd_ecommerce_accrual.BackColor = RGB(166, 166, 166)

End Sub

Private Sub cmd_internal_sales_ht_GotFocus()

Me.cmd_internal_sales_ht.BackColor = RGB(0, 176, 80)

End Sub

Private Sub cmd_internal_sales_ht_LostFocus()

Me.cmd_internal_sales_ht.BackColor = RGB(166, 166, 166)

End Sub

Private Sub cmd_internal_sales_myop_GotFocus()

Me.cmd_internal_sales_myop.BackColor = RGB(0, 176, 80)

End Sub

Private Sub cmd_internal_sales_myop_LostFocus()

Me.cmd_internal_sales_myop.BackColor = RGB(166, 166, 166)

End Sub

Private Sub cmd_reclass_po_and_toner_GotFocus()

Me.cmd_reclass_po_and_toner.BackColor = RGB(0, 176, 80)

End Sub

Private Sub cmd_reclass_po_and_toner_LostFocus()

Me.cmd_reclass_po_and_toner.BackColor = RGB(166, 166, 166)

End Sub

Private Sub cmd_reclass_tsc_kitting_cogs_GotFocus()

Me.cmd_reclass_tsc_kitting_cogs.BackColor = RGB(0, 176, 80)

End Sub

Private Sub cmd_reclass_tsc_kitting_cogs_LostFocus()

Me.cmd_reclass_tsc_kitting_cogs.BackColor = RGB(166, 166, 166)

End Sub

Private Sub cmd_update_gl_sales_cogs_GotFocus()

Me.cmd_update_gl_sales_cogs.BackColor = RGB(0, 176, 80)

End Sub

Private Sub cmd_update_gl_sales_cogs_LostFocus()

Me.cmd_update_gl_sales_cogs.BackColor = RGB(166, 166, 166)

End Sub

Private Sub cmd_comparisons_GotFocus()

Me.cmd_comparisons.BackColor = RGB(0, 176, 80)

End Sub

Private Sub cmd_comparisons_LostFocus()

Me.cmd_comparisons.BackColor = RGB(166, 166, 166)

End Sub

Private Sub cmd_sales_and_cost_total_by_gl_company_GotFocus()

Me.cmd_sales_and_cost_total_by_gl_company.BackColor = RGB(0, 176, 80)

End Sub

Private Sub cmd_sales_and_cost_total_by_gl_company_LostFocus()

Me.cmd_sales_and_cost_total_by_gl_company.BackColor = RGB(166, 166, 166)

End Sub

Private Sub cmd_montefiore_copier_GotFocus()

Me.cmd_montefiore_copier.BackColor = RGB(0, 176, 80)

End Sub

Private Sub cmd_montefiore_copier_LostFocus()

Me.cmd_montefiore_copier.BackColor = RGB(166, 166, 166)

End Sub

Private Sub cmd_line_variances_GotFocus()

Me.cmd_line_variances.BackColor = RGB(0, 176, 80)

End Sub

Private Sub cmd_line_variances_LostFocus()

Me.cmd_line_variances.BackColor = RGB(166, 166, 166)

End Sub


