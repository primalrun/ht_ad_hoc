Attribute VB_Name = "HT"
Option Compare Database



Public Sub SalesAnalysisHT()

Dim dteStart As Date
Dim dteEnd As Date
Dim qdefDAO As DAO.QueryDef

dteStart = DLookup("[value_date]", "[global]", "category='close_date_start'")
dteEnd = DLookup("[value_date]", "[global]", "category='close_date_end'")

On Error Resume Next
DoCmd.DeleteObject acTable, "sales_analysis_summary_ht"
On Error GoTo 0

strsql = ""
strsql = strsql & vbCrLf & "SELECT" & vbCrLf & _
"  ad.Company," & vbCrLf & _
"  ad.bu," & vbCrLf & _
"  sum(ad.sales) as Sales," & vbCrLf & _
"  sum(ad.cost) as Cost" & vbCrLf & _
"into sales_analysis_summary_ht" & vbCrLf & _
"from dw_sales_cost_invoice ad" & vbCrLf & _
"WHERE" & vbCrLf & _
"  ad.system = 'NAV'" & vbCrLf & _
"group BY" & vbCrLf & _
"  ad.Company," & vbCrLf & _
"  ad.bu"

CurrentDb.Execute strsql
Application.RefreshDatabaseWindow
MsgBox "Complete", vbInformation + vbSystemModal, "HT Sales Analysis"

End Sub

Public Sub ECommerceAccrual()

Dim dblSales As Double
Dim dblCost As Double
Dim dteEnd As Date
Dim qdefDAO As DAO.QueryDef
Dim appXL As Excel.Application
Dim wbCur As Excel.Workbook
Dim wbNew As Excel.Workbook
Dim wsTemp As Excel.Worksheet
Dim rsDAO As DAO.Recordset
Dim fsoTemp As FileSystemObject

dteEnd = DLookup("[value_date]", "[global]", "category='close_date_end'")

'link excel file
DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel12Xml, "temp1", varFiles(1), True
Application.RefreshDatabaseWindow

strsql = "select round(sum(t.amount),2) as Sales, round(sum(t.totcost),2) as Cost from temp1 t where t.[Qty_ Shipped Not Invoiced] <>0"
Set rsDAO = CurrentDb.OpenRecordset(strsql, dbOpenDynaset)

With rsDAO
    If .BOF = False And .EOF = False Then
        .MoveFirst
        dblSales = .Fields(0).Value
        dblCost = .Fields(1).Value
    End If
    
    .Close
End With

Set rsDAO = Nothing

strsql = ""
strsql = "select * into ecommerce_accrual from ("
strsql = strsql & vbCrLf & "select" & vbCrLf & _
"  'SJE15' as Doc_Nbr," & vbCrLf & _
"  12140 as GL_Acct," & vbCrLf & _
"  #" & dteEnd & "# as Post_Date," & vbCrLf & _
"  'E-Commerce Accrual' as Description," & vbCrLf & _
"   null as Space_1," & vbCrLf & _
"  'OFFPROD' as BU," & vbCrLf & _
"  null as Loc," & vbCrLf & _
"  null as Dept," & vbCrLf & _
"  'MYOI' as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4,  " & vbCrLf & _
"  round(sum(t.amount),2) as Amount" & vbCrLf & _
"from temp1 t" & vbCrLf & _
"where" & vbCrLf & _
"  t.[Qty_ Shipped Not Invoiced] <>0" & vbCrLf & _
"" & vbCrLf & _
"union all" & vbCrLf & _
""
strsql = strsql & vbCrLf & "select" & vbCrLf & _
"  'SJE15' as Doc_Nbr," & vbCrLf & _
"  35100 as GL_Acct," & vbCrLf & _
"  #" & dteEnd & "# as Post_Date," & vbCrLf & _
"  'E-Commerce Accrual' as Description," & vbCrLf & _
"   null as Space_1," & vbCrLf & _
"  'OFFPROD' as BU," & vbCrLf & _
"  null as Loc," & vbCrLf & _
"  null as Dept," & vbCrLf & _
"  'MYOI' as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4,  " & vbCrLf & _
"  round(sum(t.amount),2)*-1 as Amount" & vbCrLf & _
"from temp1 t" & vbCrLf & _
"where" & vbCrLf & _
"  t.[Qty_ Shipped Not Invoiced] <>0" & vbCrLf & _
"" & vbCrLf & _
"union all" & vbCrLf & _
""
strsql = strsql & vbCrLf & "select" & vbCrLf & _
"  'SJE15' as Doc_Nbr," & vbCrLf & _
"  45100 as GL_Acct," & vbCrLf & _
"  #" & dteEnd & "# as Post_Date," & vbCrLf & _
"  'E-Commerce Accrual' as Description," & vbCrLf & _
"   null as Space_1," & vbCrLf & _
"  'OFFPROD' as BU," & vbCrLf & _
"  null as Loc," & vbCrLf & _
"  null as Dept," & vbCrLf & _
"  'MYOI' as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4,  " & vbCrLf & _
"  round(sum(t.totcost),2) as Amount" & vbCrLf & _
"from temp1 t" & vbCrLf & _
"where" & vbCrLf & _
"  t.[Qty_ Shipped Not Invoiced] <>0" & vbCrLf & _
"" & vbCrLf & _
"union all" & vbCrLf & _
""
strsql = strsql & vbCrLf & "select" & vbCrLf & _
"  'SJE15' as Doc_Nbr," & vbCrLf & _
"  23310 as GL_Acct," & vbCrLf & _
"  #" & dteEnd & "# as Post_Date," & vbCrLf & _
"  'E-Commerce Accrual' as Description," & vbCrLf & _
"   null as Space_1," & vbCrLf & _
"  'OFFPROD' as BU," & vbCrLf & _
"  null as Loc," & vbCrLf & _
"  null as Dept," & vbCrLf & _
"  'MYOI' as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4,  " & vbCrLf & _
"  round(sum(t.totcost),2)*-1 as Amount" & vbCrLf & _
"from temp1 t" & vbCrLf & _
"where" & vbCrLf & _
"  t.[Qty_ Shipped Not Invoiced] <>0" & vbCrLf & _
"" & vbCrLf & _
"union all" & vbCrLf & _
""
strsql = strsql & vbCrLf & "select" & vbCrLf & _
"  'SJE15' as Doc_Nbr," & vbCrLf & _
"  12140 as GL_Acct," & vbCrLf & _
"  #" & DateAdd("d", 1, dteEnd) & "# as Post_Date," & vbCrLf & _
"  'E-Commerce Accrual' as Description," & vbCrLf & _
"   null as Space_1," & vbCrLf & _
"  'OFFPROD' as BU," & vbCrLf & _
"  null as Loc," & vbCrLf & _
"  null as Dept," & vbCrLf & _
"  'MYOI' as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4,  " & vbCrLf & _
"  round(sum(t.amount),2)*-1 as Amount" & vbCrLf & _
"from temp1 t" & vbCrLf & _
"where" & vbCrLf & _
"  t.[Qty_ Shipped Not Invoiced] <>0" & vbCrLf & _
"" & vbCrLf & _
"union all" & vbCrLf & _
""
strsql = strsql & vbCrLf & "select" & vbCrLf & _
"  'SJE15' as Doc_Nbr," & vbCrLf & _
"  35100 as GL_Acct," & vbCrLf & _
"  #" & DateAdd("d", 1, dteEnd) & "# as Post_Date," & vbCrLf & _
"  'E-Commerce Accrual' as Description," & vbCrLf & _
"   null as Space_1," & vbCrLf & _
"  'OFFPROD' as BU," & vbCrLf & _
"  null as Loc," & vbCrLf & _
"  null as Dept," & vbCrLf & _
"  'MYOI' as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4,  " & vbCrLf & _
"  round(sum(t.amount),2) as Amount" & vbCrLf & _
"from temp1 t" & vbCrLf & _
"where" & vbCrLf & _
"  t.[Qty_ Shipped Not Invoiced] <>0" & vbCrLf & _
"" & vbCrLf & _
"union all" & vbCrLf & _
""
strsql = strsql & vbCrLf & "select" & vbCrLf & _
"  'SJE15' as Doc_Nbr," & vbCrLf & _
"  45100 as GL_Acct," & vbCrLf & _
"  #" & DateAdd("d", 1, dteEnd) & "# as Post_Date," & vbCrLf & _
"  'E-Commerce Accrual' as Description," & vbCrLf & _
"   null as Space_1," & vbCrLf & _
"  'OFFPROD' as BU," & vbCrLf & _
"  null as Loc," & vbCrLf & _
"  null as Dept," & vbCrLf & _
"  'MYOI' as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4,  " & vbCrLf & _
"  round(sum(t.totcost),2)*-1 as Amount" & vbCrLf & _
"from temp1 t" & vbCrLf & _
"where" & vbCrLf & _
"  t.[Qty_ Shipped Not Invoiced] <>0" & vbCrLf & _
"" & vbCrLf & _
"union all" & vbCrLf & _
""
strsql = strsql & vbCrLf & "select" & vbCrLf & _
"  'SJE15' as Doc_Nbr," & vbCrLf & _
"  23310 as GL_Acct," & vbCrLf & _
"  #" & DateAdd("d", 1, dteEnd) & "# as Post_Date," & vbCrLf & _
"  'E-Commerce Accrual' as Description," & vbCrLf & _
"   null as Space_1," & vbCrLf & _
"  'OFFPROD' as BU," & vbCrLf & _
"  null as Loc," & vbCrLf & _
"  null as Dept," & vbCrLf & _
"  'MYOI' as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4,  " & vbCrLf & _
"  round(sum(t.totcost),2) as Amount" & vbCrLf & _
"from temp1 t" & vbCrLf & _
"where" & vbCrLf & _
"  t.[Qty_ Shipped Not Invoiced] <>0"

strsql = strsql & vbCrLf & ") q1"

On Error Resume Next
DoCmd.DeleteObject acTable, "ecommerce_accrual"
On Error GoTo 0

CurrentDb.Execute strsql
DoCmd.DeleteObject acTable, "temp1"
Application.RefreshDatabaseWindow

Set fsoTemp = New FileSystemObject

If fsoTemp.FileExists(strPathSave & "E-Commerce Accrual.xlsx") = True Then
    fsoTemp.DeleteFile strPathSave & "E-Commerce Accrual.xlsx", True
End If

Set fsoTemp = Nothing

DoCmd.TransferSpreadsheet transfertype:=acExport, _
    spreadsheettype:=acSpreadsheetTypeExcel12Xml, _
    TableName:="ecommerce_accrual", _
    FileName:=strPathSave & "E-Commerce Accrual.xlsx", _
    hasfieldnames:=True
    
Set fsoTemp = New FileSystemObject

If fsoTemp.FileExists("\\tndcfile05\Departments\Accounting\Accounting\JE_Uploads\" & "E-Commerce Accrual.csv") = True Then
    fsoTemp.DeleteFile "\\tndcfile05\Departments\Accounting\Accounting\JE_Uploads\" & "E-Commerce Accrual.csv", True
End If

Set fsoTemp = Nothing

Set appXL = New Excel.Application
appXL.Visible = True
Set wbCur = appXL.Workbooks.Open(strPathSave & "E-Commerce Accrual.xlsx")
Set wbNew = appXL.Workbooks.Add
wbCur.Activate
wbCur.Sheets(1).Select
wbCur.Sheets("ecommerce_accrual").Copy before:=wbNew.Sheets(1)
wbNew.Activate
appXL.ActiveSheet.Name = "upload"

For Each wsTemp In appXL.Worksheets
    If wsTemp.Name <> "upload" Then
        appXL.Sheets(wsTemp.Name).Delete
    End If
Next wsTemp

appXL.Rows(1).Delete
wbNew.SaveAs "\\tndcfile05\Departments\Accounting\Accounting\JE_Uploads\" & "E-Commerce Accrual.csv", xlCSV
wbNew.Close False
wbCur.Sheets("ecommerce_accrual").Activate
wbCur.Save
wbCur.Close True
appXL.Quit
Set appXL = Nothing
    
CurrentDb.Execute "delete from ecommerce_accrual_trxn where post_date = #" & dteEnd & "#"
CurrentDb.Execute "insert into ecommerce_accrual_trxn (post_date, sales, cost) values(#" & dteEnd & "#, " & dblSales & ", " & dblCost & ")"
    
End Sub

Public Sub InternalSalesHitouch()

'variables
Dim intRowLast As Integer
Dim intRowNext As Integer
Dim intRowTemp As Integer
Dim intRows As Integer
Dim dblTax As Double
Dim dteStart As Date
Dim dteEnd As Date
Dim strFile As String
Dim strCategory As String
Dim cmdadodb As ADODB.Command
Dim wbCur As Excel.Workbook
Dim wbNew As Excel.Workbook
Dim wsJenniferPrintout As Excel.Worksheet
Dim fsoTemp As FileSystemObject
Dim qdefDAO As DAO.QueryDef
Dim rsDAO As DAO.Recordset

dteStart = DLookup("[value_date]", "[global]", "category='close_date_start'")
dteEnd = DLookup("[value_date]", "[global]", "category='close_date_end'")

'import ht data
Call PublicStuff.GetFile("Select Invoice Register Finance File")

If blnSelectFile = False Then
    MsgBox "No file selected, process cancelled", vbCritical + vbSystemModal, "File Validation"
    Exit Sub
End If

'link excel file
On Error Resume Next
DoCmd.DeleteObject acTable, "temp1"
On Error GoTo 0

DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel12Xml, "temp1", varFiles(1), True
Application.RefreshDatabaseWindow

'drop/create internal_sales_ht
On Error Resume Next
DoCmd.DeleteObject acTable, "internal_sales_ht"
On Error GoTo 0

CurrentDb.Execute "select * into internal_sales_ht from temp1 t where t.salesperson = '1017 - INTERNAL SALES'"

'drop linked table
DoCmd.DeleteObject acTable, "temp1"

'update null dept
CurrentDb.Execute "update internal_sales_ht set dept = customername where dept = 'NULL'"

'delete records where dept = 'EMPLOYEE PURCHASE ACCOUNT'
CurrentDb.Execute "delete from internal_sales_ht where dept = 'EMPLOYEE PURCHASE ACCOUNT'"

'add type column
Set cmdadodb = New ADODB.Command
With cmdadodb
    .ActiveConnection = CurrentProject.Connection
    .CommandType = adCmdText
    .CommandText = "alter table internal_sales_ht add column type varchar(50)"
    .Execute
End With
Set cmdadodb = Nothing

'update type
CurrentDb.Execute "update internal_sales_ht i left join dept_type d on i.dept = d.dept set i.type = d.type"

'export data sets
Call PublicStuff.SelectSaveDirectory("Select Directory to Save HT Internal Sales Data")

If strPathSave = "" Then
    MsgBox "No Save Directory Selected, Process Cancelled", vbCritical + vbSystemModal, "Save Directory Validation"
    Exit Sub
End If

On Error Resume Next
DoCmd.DeleteObject acTable, "Jennifer_Printout"
On Error GoTo 0

strsql = ""
strsql = strsql & vbCrLf & "select" & vbCrLf & _
"  h.id as ID," & vbCrLf & _
"  i.dept as Dept," & vbCrLf & _
"  i.type as Type," & vbCrLf & _
"  sum(i.amount) as Sales," & vbCrLf & _
"  sum(i.unloadedcost) as Cost" & vbCrLf & _
"into Jennifer_Printout " & vbCrLf & _
"from internal_sales_ht i" & vbCrLf & _
"  inner join type_sort_ht h" & vbCrLf & _
"    on i.type = h.type" & vbCrLf & _
"group by" & vbCrLf & _
"  h.id," & vbCrLf & _
"  i.dept," & vbCrLf & _
"  i.type" & vbCrLf & _
"order by 1, 2"

CurrentDb.Execute strsql
Application.RefreshDatabaseWindow

strFile = "Internal Sales HT " & Format(dteStart, "YYYYMM") & ".xlsx"

Set fsoTemp = New FileSystemObject

If fsoTemp.FileExists(strPathSave & strFile) = True Then
    fsoTemp.DeleteFile strPathSave & strFile, True
End If

Set fsoTemp = Nothing

On Error Resume Next
DoCmd.DeleteObject acTable, "rev_cogs_ic_ht"
On Error GoTo 0

'JE
strsql = ""
strsql = strsql & vbCrLf & "set NOCOUNT on" & vbCrLf & _
"DECLARE @datestart date = '" & CStr(dteStart) & "'" & vbCrLf & _
"declare @dateend date = '" & CStr(dteEnd) & "'" & vbCrLf & _
"" & vbCrLf & _
"SELECT" & vbCrLf & _
"  -sum(case when gl.[G_L Account No_] = '32100' then gl.Amount else 0 end) as Sales," & vbCrLf & _
"  sum(case when gl.[G_L Account No_] = '42100' then gl.Amount else 0 end) as Cost" & vbCrLf & _
"from navrep.dbo.[Hi Touch$G_L Entry] gl with(nolock)" & vbCrLf & _
"WHERE" & vbCrLf & _
"  gl.[G_L Account No_] in ('32100', '42100')" & vbCrLf & _
"  and gl.[Posting Date] BETWEEN @datestart and @dateend" & vbCrLf & _
"  and gl.[Source Code] <> 'GENJNL'"

Set qdefDAO = CurrentDb.QueryDefs("q_navrep_select")
qdefDAO.SQL = strsql
Set qdefDAO = Nothing
CurrentDb.Execute "select * into rev_cogs_ic_ht from q_navrep_select"

On Error Resume Next
DoCmd.DeleteObject acTable, "je"
On Error GoTo 0

strsql = ""
strsql = strsql & vbCrLf & "select" & vbCrLf & _
"  *" & vbCrLf & _
"into je" & vbCrLf & _
"from (" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  'SJE21' as Doc_Nbr," & vbCrLf & _
"  32100 as GL_Acct," & vbCrLf & _
"  #" & dteEnd & "# as Post_Date," & vbCrLf & _
"  'Write-off HT I/C Revenue' as Description," & vbCrLf & _
"  null as Space_1," & vbCrLf & _
"  'OFFPROD' as BU," & vbCrLf & _
"  null as Loc," & vbCrLf & _
"  null as Dept," & vbCrLf & _
"  'HT' as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4," & vbCrLf & _
"  round(rc.sales,2) as Amount" & vbCrLf & _
"from rev_cogs_ic_ht rc"
strsql = strsql & vbCrLf & "" & vbCrLf & _
"union all" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  'SJE21' as Doc_Nbr," & vbCrLf & _
"  42100 as GL_Acct," & vbCrLf & _
"  #" & dteEnd & "# as Post_Date," & vbCrLf & _
"  'Write-off HT I/C COGS' as Description," & vbCrLf & _
"  null as Space_1," & vbCrLf & _
"  'OFFPROD' as BU," & vbCrLf & _
"  null as Loc," & vbCrLf & _
"  null as Dept," & vbCrLf & _
"  'HT' as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4," & vbCrLf & _
"  -round(rc.cost,2) as Amount" & vbCrLf & _
"from rev_cogs_ic_ht rc" & vbCrLf & _
"" & vbCrLf & _
"union all"
strsql = strsql & vbCrLf & "" & vbCrLf & _
"select" & vbCrLf & _
"  'SJE21' as Doc_Nbr," & vbCrLf & _
"  45100 as GL_Acct," & vbCrLf & _
"  #" & dteEnd & "# as Post_Date," & vbCrLf & _
"  'Write-off HT I/C COGS' as Description," & vbCrLf & _
"  null as Space_1," & vbCrLf & _
"  'OFFPROD' as BU," & vbCrLf & _
"  null as Loc," & vbCrLf & _
"  null as Dept," & vbCrLf & _
"  'HT' as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4," & vbCrLf & _
"  -round((select sum(cost) from jennifer_printout) - rc.cost,2) as Amount" & vbCrLf & _
"from rev_cogs_ic_ht rc" & vbCrLf & _
"" & vbCrLf & _
"union all" & vbCrLf & _
"" & vbCrLf & _
"select"
strsql = strsql & vbCrLf & "  'SJE21' as Doc_Nbr," & vbCrLf & _
"  17100 as GL_Acct," & vbCrLf & _
"  #" & dteEnd & "# as Post_Date," & vbCrLf & _
"  'MRO Supplies RAC' as Description," & vbCrLf & _
"  null as Space_1," & vbCrLf & _
"  'RAC' as BU," & vbCrLf & _
"  null as Loc," & vbCrLf & _
"  null as Dept," & vbCrLf & _
"  'RAC' as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4," & vbCrLf & _
"  round(sum(jp.cost),2) as Amount" & vbCrLf & _
"from jennifer_printout jp" & vbCrLf & _
"where" & vbCrLf & _
"  jp.type = 'Due from RAC'" & vbCrLf & _
"" & vbCrLf & _
"union all" & vbCrLf & _
"" & vbCrLf & _
"select"
strsql = strsql & vbCrLf & "  'SJE21' as Doc_Nbr," & vbCrLf & _
"  41150 as GL_Acct," & vbCrLf & _
"  #" & dteEnd & "# as Post_Date," & vbCrLf & _
"  'MRO Supplies Warehouse' as Description," & vbCrLf & _
"  null as Space_1," & vbCrLf & _
"  'OFFPROD' as BU," & vbCrLf & _
"  null as Loc," & vbCrLf & _
"  null as Dept," & vbCrLf & _
"  'HT' as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4," & vbCrLf & _
"  round(sum(jp.cost),2) as Amount" & vbCrLf & _
"from jennifer_printout jp" & vbCrLf & _
"where" & vbCrLf & _
"  jp.dept = '1002-WAREHOUSE - NJ'" & vbCrLf & _
"" & vbCrLf & _
"union all" & vbCrLf & _
"" & vbCrLf & _
"select"
strsql = strsql & vbCrLf & "  'SJE21' as Doc_Nbr," & vbCrLf & _
"  55670 as GL_Acct," & vbCrLf & _
"  #" & dteEnd & "# as Post_Date," & vbCrLf & _
"  'MRO Supplies Office' as Description," & vbCrLf & _
"  null as Space_1," & vbCrLf & _
"  'OFFPROD' as BU," & vbCrLf & _
"  null as Loc," & vbCrLf & _
"  null as Dept," & vbCrLf & _
"  'HT' as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4," & vbCrLf & _
"  round(sum(jp.cost),2) as Amount" & vbCrLf & _
"from jennifer_printout jp" & vbCrLf & _
"where" & vbCrLf & _
"  jp.type = 'MRO Supplies'" & vbCrLf & _
"  and jp.dept <> '1002-WAREHOUSE - NJ'" & vbCrLf & _
"" & vbCrLf & _
"union all" & vbCrLf & _
""
strsql = strsql & vbCrLf & "select" & vbCrLf & _
"  'SJE21' as Doc_Nbr," & vbCrLf & _
"  52200 as GL_Acct," & vbCrLf & _
"  #" & dteEnd & "# as Post_Date," & vbCrLf & _
"  'Marketing Expense' as Description," & vbCrLf & _
"  null as Space_1," & vbCrLf & _
"  'OFFPROD' as BU," & vbCrLf & _
"  null as Loc," & vbCrLf & _
"  null as Dept," & vbCrLf & _
"  'HT' as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4," & vbCrLf & _
"  round(sum(jp.cost),2) as Amount" & vbCrLf & _
"from jennifer_printout jp" & vbCrLf & _
"where" & vbCrLf & _
"  jp.type = 'Sales/Marketing'" & vbCrLf & _
"" & vbCrLf & _
"union all" & vbCrLf & _
""
strsql = strsql & vbCrLf & "select" & vbCrLf & _
"  'SJE21' as Doc_Nbr," & vbCrLf & _
"  54100 as GL_Acct," & vbCrLf & _
"  #" & dteEnd & "# as Post_Date," & vbCrLf & _
"  'IT Expense' as Description," & vbCrLf & _
"  null as Space_1," & vbCrLf & _
"  'OFFPROD' as BU," & vbCrLf & _
"  null as Loc," & vbCrLf & _
"  null as Dept," & vbCrLf & _
"  'HT' as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4," & vbCrLf & _
"  round(sum(jp.cost),2) as Amount" & vbCrLf & _
"from jennifer_printout jp" & vbCrLf & _
"where" & vbCrLf & _
"  jp.type = 'IT Expense') q1"

CurrentDb.Execute strsql
Application.RefreshDatabaseWindow

Set rsDAO = CurrentDb.OpenRecordset("select sum(amount) from je", dbOpenDynaset)

If rsDAO.BOF = False And rsDAO.EOF = False Then
    dblTax = -round(rsDAO.Fields(0).Value, 2)
End If

rsDAO.Close
Set rsDAO = Nothing

CurrentDb.Execute "insert into je (doc_nbr, gl_acct, post_date, description, amount) values('SJE21', 22100, #" & dteEnd & "#, 'Sales Tax Payable', " & dblTax & ")"

DoCmd.TransferSpreadsheet transfertype:=acExport, spreadsheettype:=acSpreadsheetTypeExcel12Xml, _
    TableName:="Jennifer_Printout", _
    FileName:=strPathSave & strFile, _
    hasfieldnames:=True
    
DoCmd.TransferSpreadsheet transfertype:=acExport, spreadsheettype:=acSpreadsheetTypeExcel12Xml, _
    TableName:="internal_sales_ht", _
    FileName:=strPathSave & strFile, _
    hasfieldnames:=True
    
DoCmd.TransferSpreadsheet transfertype:=acExport, spreadsheettype:=acSpreadsheetTypeExcel12Xml, _
    TableName:="je", _
    FileName:=strPathSave & strFile, _
    hasfieldnames:=True
    
Set appXL = New Excel.Application
Set wbCur = appXL.Workbooks.Open(strPathSave & strFile)
Set wsJenniferPrintout = wbCur.Sheets("Jennifer_Printout")
wsJenniferPrintout.Activate
appXL.Columns(1).Delete
appXL.Rows(1).Delete
intRowLast = appXL.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
intRowTemp = intRowLast

Do Until intRowTemp = 1
    intRow = intRowTemp - 1
    If appXL.Cells(intRowTemp, 2).Value <> appXL.Cells(intRow, 2).Value Then
        appXL.Cells(intRowTemp, 2).EntireRow.Insert
        appXL.Cells(intRowTemp, 2).EntireRow.Insert
    End If
    
    intRowTemp = intRowTemp - 1
Loop

intRowLast = appXL.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
intRowTemp = intRowLast + 1

Do Until intRowTemp = 1
    intRow = intRowTemp - 1
    
    If appXL.Cells(intRow, 2).Value <> "" And appXL.Cells(intRowTemp, 2).Value = "" Then
        appXL.Cells(intRowTemp, 1).Value = appXL.Cells(intRow, 2).Value & " Total"
        appXL.Cells(intRowTemp, 1).Font.Bold = True
        intRows = appXL.Cells(intRow, 2).CurrentRegion.Rows.Count - 1
        appXL.Cells(intRowTemp, 3).FormulaR1C1 = "=sum(r[-" & intRows & "]c:r[-1]c)"
        appXL.Cells(intRowTemp, 4).FormulaR1C1 = "=sum(r[-" & intRows & "]c:r[-1]c)"
        appXL.Range(appXL.Cells(intRowTemp, 3), appXL.Cells(intRowTemp, 4)).Borders(xlEdgeTop).Weight = xlThin
    End If
    
    intRowTemp = intRowTemp - 1
Loop



appXL.Cells(1, 2).EntireColumn.Delete
intRowLast = appXL.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
appXL.Range(appXL.Cells(1, 2), appXL.Cells(intRowLast, 3)).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"

appXL.Cells(1, 1).EntireRow.Insert
appXL.Cells(1, 1).EntireRow.Insert
appXL.Cells(1, 1).EntireRow.Insert

appXL.Cells(1, 1).Value = Format(dteStart, "MMMM YYYY")
appXL.Cells(1, 1).Font.Bold = True
appXL.Range(appXL.Cells(1, 1), appXL.Cells(1, 3)).HorizontalAlignment = xlCenterAcrossSelection

With appXL.Cells(3, 1)
    .Value = "Department"
    .Font.Bold = True
    .HorizontalAlignment = xlLeft
End With

With appXL.Cells(3, 2)
    .Value = "Sales"
    .Font.Bold = True
    .HorizontalAlignment = xlRight
End With

With appXL.Cells(3, 3)
    .Value = "Cost"
    .Font.Bold = True
    .HorizontalAlignment = xlRight
End With

intRowLast = appXL.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
intRow = intRowLast + 2

appXL.Cells(intRow, 2).FormulaR1C1 = "=sumif(r4c1:r[-1]c1," & Chr(34) & "*Total*" & Chr(34) & ",r4c:r[-1]c)"
appXL.Cells(intRow, 3).FormulaR1C1 = "=sumif(r4c1:r[-1]c1," & Chr(34) & "*Total*" & Chr(34) & ",r4c:r[-1]c)"
appXL.Cells(intRow, 1).Value = "Total"
appXL.Cells(intRow, 1).Font.Bold = True
appXL.Range(appXL.Cells(intRow, 2), appXL.Cells(intRow, 3)).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"

appXL.Columns(1).AutoFit
appXL.Range(appXL.Cells(1, 2), appXL.Cells(1, 3)).EntireColumn.ColumnWidth = 13
Call PageSetupPortrait1
appXL.Cells(1, 1).Select

Set wbNew = appXL.Workbooks.Add
wbCur.Activate
wbCur.Sheets("je").Copy before:=wbNew.Sheets(1)
wbNew.Activate
appXL.ActiveSheet.Name = "upload"

For Each wsTemp In appXL.Worksheets
    If wsTemp.Name <> "upload" Then
        appXL.Sheets(wsTemp.Name).Delete
    End If
Next wsTemp

Set fsoTemp = New FileSystemObject

If fsoTemp.FileExists("\\tndcfile05\Departments\Accounting\Accounting\JE_Uploads\" & "Internal Sales HT.csv") = True Then
    fsoTemp.DeleteFile "\\tndcfile05\Departments\Accounting\Accounting\JE_Uploads\" & "Internal Sales HT.csv", True
End If

Set fsoTemp = Nothing

appXL.Rows(1).Delete
wbNew.SaveAs "\\tndcfile05\Departments\Accounting\Accounting\JE_Uploads\" & "Internal Sales HT.csv", xlCSV
wbNew.Close False

wbCur.Save
wbCur.Close False
appXL.Quit
Set appXL = Nothing


MsgBox "Complete", vbInformation + vbSystemModal, "HT Internal Sales"


End Sub

Private Sub PageSetupPortrait1()

Dim lngRowLast As Long
Dim lngColLast As Long
Dim rngPrint As Range

lngRowLast = appXL.Cells.Find(What:="*", After:=appXL.Cells(1, 1), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
lngColLast = appXL.Cells.Find(What:="*", After:=appXL.Cells(1, 1), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
Set rngPrint = appXL.Cells(lngRowLast, lngColLast)

With appXL.ActiveSheet.PageSetup
        .PrintArea = appXL.Cells(1, 1).Address(False, False) & ":" & rngPrint.Address(False, False)
        .PrintTitleRows = "$1:$3"
        .LeftFooter = "Corporate Finance"
        .CenterFooter = "&Z" & Chr(10) & "&F"
        .RightFooter = "Page &P of &N"
        .LeftMargin = appXL.InchesToPoints(0.3)
        .RightMargin = appXL.InchesToPoints(0.3)
        .TopMargin = appXL.InchesToPoints(0.5)
        .BottomMargin = appXL.InchesToPoints(0.75)
        .HeaderMargin = appXL.InchesToPoints(0.3)
        .FooterMargin = appXL.InchesToPoints(0.3)
        .CenterHorizontally = True
        .Orientation = xlPortrait
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 2000
End With

End Sub

Public Sub ReclassPOandToner()

Dim dteEnd As Date
Dim strTonerName As String
Dim strPOName As String
Dim strAmount As String
Dim strCost As String
Dim strFile As String
Dim wbCur As Excel.Workbook
Dim wbNew As Excel.Workbook
Dim db As DAO.Database
Dim tdef As DAO.TableDef
Dim fld As DAO.Field
Dim fsoTemp As FileSystemObject

Set db = CurrentDb
dteEnd = DLookup("[value_date]", "[global]", "category='close_date_end'")

On Error Resume Next
DoCmd.DeleteObject acTable, "toner"
DoCmd.DeleteObject acTable, "po"
DoCmd.DeleteObject acTable, "po_toner_reclass"
DoCmd.DeleteObject acTable, "je"
On Error GoTo 0

strTonerName = InputBox("Toner Sheet Name", "Toner Sheet Name")
strPOName = InputBox("PO Sheet Name", "PO Sheet Name")
strTonerName = strTonerName & "!"
strPOName = strPOName & "!"

If Len(strTonerName) = 0 Or Len(strPOName) = 0 Then
    MsgBox "Sheet Names not provided, Process Cancelled", vbCritical + vbSystemModal, "Sheet Name"
    Exit Sub
End If
    
'link Toner
DoCmd.TransferSpreadsheet transfertype:=acLink, _
    spreadsheettype:=acSpreadsheetTypeExcel12Xml, _
    TableName:="toner", _
    FileName:=varFiles(1), _
    hasfieldnames:=True, _
    Range:=strTonerName
    
DoCmd.TransferSpreadsheet transfertype:=acLink, _
    spreadsheettype:=acSpreadsheetTypeExcel12Xml, _
    TableName:="po", _
    FileName:=varFiles(1), _
    hasfieldnames:=True, _
    Range:=strPOName
    
Application.RefreshDatabaseWindow

For Each fld In db.TableDefs("toner").Fields
    If fld.Name Like "*Amount*" Then
        strAmount = fld.Name
    End If
    If fld.Name Like "*Cost*" Then
        strCost = fld.Name
    End If
Next fld



strsql = ""
strsql = strsql & vbCrLf & "select" & vbCrLf & _
"  q1.invoice," & vbCrLf & _
"  q1.bu," & vbCrLf & _
"  dw1.bu as bu_dw," & vbCrLf & _
"  sum(q1.sales) as sales," & vbCrLf & _
"  sum(q1.cost) as cost" & vbCrLf & _
"into po_toner_reclass" & vbCrLf & _
"from (" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  t.invoicenbr as invoice," & vbCrLf & _
"  t.itembusunit as bu," & vbCrLf & _
"  sum(t.[" & strAmount & "]) as sales," & vbCrLf & _
"  sum(t.[" & strCost & "]) as cost" & vbCrLf & _
"from toner t" & vbCrLf & _
"group by" & vbCrLf & _
"  t.invoicenbr," & vbCrLf & _
"  t.itembusunit" & vbCrLf & _
"" & vbCrLf & _
"union all"
strsql = strsql & vbCrLf & "" & vbCrLf & _
"select" & vbCrLf & _
"  p.inv_no," & vbCrLf & _
"  p.busunit," & vbCrLf & _
"  sum(p.amount) as sales," & vbCrLf & _
"  sum(p.cost) as cost" & vbCrLf & _
"from po p" & vbCrLf & _
"group by" & vbCrLf & _
"  p.inv_no," & vbCrLf & _
"  p.busunit) q1" & vbCrLf & _
"" & vbCrLf & _
"  left join (select dw.invoice, max(dw.bu) as bu from dw_sales_cost_invoice dw group by invoice) dw1" & vbCrLf & _
"    on q1.invoice = dw1.invoice" & vbCrLf & _
"" & vbCrLf & _
"group by" & vbCrLf & _
"  q1.invoice," & vbCrLf & _
"  q1.bu," & vbCrLf & _
"  dw1.bu"

db.Execute strsql

'je
strsql = ""
strsql = strsql & vbCrLf & "select * into je from (" & vbCrLf & _
"select" & vbCrLf & _
"  'SJE17' as Doc_Nbr," & vbCrLf & _
"  35100 as GL_Acct," & vbCrLf & _
"  #" & dteEnd & "# as Post_Date," & vbCrLf & _
"  'Reclass PO and Toner Sales and COGS' as Description," & vbCrLf & _
"  null as Space_1," & vbCrLf & _
"  'OFFPROD' as BU," & vbCrLf & _
"  null as Loc," & vbCrLf & _
"  null as Dept," & vbCrLf & _
"  null as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4," & vbCrLf & _
"  round(sum(p.sales),2) as Amount" & vbCrLf & _
"from po_toner_reclass p" & vbCrLf & _
"where" & vbCrLf & _
"  p.bu_dw = 'Office Products'" & vbCrLf & _
"group by" & vbCrLf & _
"  'SJE17',"
strsql = strsql & vbCrLf & "  35100," & vbCrLf & _
"  #" & dteEnd & "#," & vbCrLf & _
"  'Reclass PO and Toner Sales and COGS'," & vbCrLf & _
"  null," & vbCrLf & _
"  'OFFPROD'," & vbCrLf & _
"  null," & vbCrLf & _
"  null," & vbCrLf & _
"  null," & vbCrLf & _
"  null," & vbCrLf & _
"  null," & vbCrLf & _
"  null" & vbCrLf & _
"" & vbCrLf & _
"union all" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  'SJE17' as Doc_Nbr," & vbCrLf & _
"  34000 as GL_Acct," & vbCrLf & _
"  #" & dteEnd & "# as Post_Date," & vbCrLf & _
"  'Reclass PO and Toner Sales and COGS' as Description," & vbCrLf & _
"  null as Space_1,"
strsql = strsql & vbCrLf & "  'MPS' as BU," & vbCrLf & _
"  null as Loc," & vbCrLf & _
"  null as Dept," & vbCrLf & _
"  null as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4," & vbCrLf & _
"  -round(sum(p.sales),2) as Amount" & vbCrLf & _
"from po_toner_reclass p" & vbCrLf & _
"where" & vbCrLf & _
"  p.bu_dw = 'Office Products'" & vbCrLf & _
"group by" & vbCrLf & _
"  'SJE17'," & vbCrLf & _
"  34000," & vbCrLf & _
"  #" & dteEnd & "#," & vbCrLf & _
"  'Reclass PO and Toner Sales and COGS'," & vbCrLf & _
"  null," & vbCrLf & _
"  'MPS'," & vbCrLf & _
"  null," & vbCrLf & _
"  null,"
strsql = strsql & vbCrLf & "  null," & vbCrLf & _
"  null," & vbCrLf & _
"  null," & vbCrLf & _
"  null" & vbCrLf & _
"" & vbCrLf & _
"union all" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  'SJE17' as Doc_Nbr," & vbCrLf & _
"  45100 as GL_Acct," & vbCrLf & _
"  #" & dteEnd & "# as Post_Date," & vbCrLf & _
"  'Reclass PO and Toner Sales and COGS' as Description," & vbCrLf & _
"  null as Space_1," & vbCrLf & _
"  'OFFPROD' as BU," & vbCrLf & _
"  null as Loc," & vbCrLf & _
"  null as Dept," & vbCrLf & _
"  null as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4,"
strsql = strsql & vbCrLf & "  -round(sum(p.cost),2) as Amount" & vbCrLf & _
"from po_toner_reclass p" & vbCrLf & _
"where" & vbCrLf & _
"  p.bu_dw = 'Office Products'" & vbCrLf & _
"group by" & vbCrLf & _
"  'SJE17'," & vbCrLf & _
"  45100," & vbCrLf & _
"  #" & dteEnd & "#," & vbCrLf & _
"  'Reclass PO and Toner Sales and COGS'," & vbCrLf & _
"  null," & vbCrLf & _
"  'OFFPROD'," & vbCrLf & _
"  null," & vbCrLf & _
"  null," & vbCrLf & _
"  null," & vbCrLf & _
"  null," & vbCrLf & _
"  null," & vbCrLf & _
"  null" & vbCrLf & _
"" & vbCrLf & _
"union all" & vbCrLf & _
""
strsql = strsql & vbCrLf & "select" & vbCrLf & _
"  'SJE17' as Doc_Nbr," & vbCrLf & _
"  44000 as GL_Acct," & vbCrLf & _
"  #" & dteEnd & "# as Post_Date," & vbCrLf & _
"  'Reclass PO and Toner Sales and COGS' as Description," & vbCrLf & _
"  null as Space_1," & vbCrLf & _
"  'MPS' as BU," & vbCrLf & _
"  null as Loc," & vbCrLf & _
"  null as Dept," & vbCrLf & _
"  null as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4," & vbCrLf & _
"  round(sum(p.cost),2) as Amount" & vbCrLf & _
"from po_toner_reclass p" & vbCrLf & _
"where" & vbCrLf & _
"  p.bu_dw = 'Office Products'" & vbCrLf & _
"group by" & vbCrLf & _
"  'SJE17'," & vbCrLf & _
"  44000,"
strsql = strsql & vbCrLf & "  #" & dteEnd & "#," & vbCrLf & _
"  'Reclass PO and Toner Sales and COGS'," & vbCrLf & _
"  null," & vbCrLf & _
"  'MPS'," & vbCrLf & _
"  null," & vbCrLf & _
"  null," & vbCrLf & _
"  null," & vbCrLf & _
"  null," & vbCrLf & _
"  null," & vbCrLf & _
"  null) q1"

db.Execute strsql

'Save File
Call PublicStuff.SelectSaveDirectory("Select Directory to Save PO and Toner Reclass")

If strPathSave = "" Then
    MsgBox "No Save Directory Selected, Process Cancelled", vbCritical + vbSystemModal, "Save Directory Validation"
    Exit Sub
End If

strFile = "PO and Toner Reclass " & Format(dteStart, "YYYYMM") & ".xlsx"

Set fsoTemp = New FileSystemObject

If fsoTemp.FileExists(strPathSave & strFile) = True Then
    fsoTemp.DeleteFile strPathSave & strFile, True
End If

If fsoTemp.FileExists("\\tndcfile05\Departments\Accounting\Accounting\JE_Uploads\" & "PO and Toner Reclass.csv") = True Then
    fsoTemp.DeleteFile "\\tndcfile05\Departments\Accounting\Accounting\JE_Uploads\" & "PO and Toner Reclass.csv", True
End If

Set fsoTemp = Nothing

DoCmd.TransferSpreadsheet transfertype:=acExport, spreadsheettype:=acSpreadsheetTypeExcel12Xml, _
    TableName:="po_toner_reclass", _
    FileName:=strPathSave & strFile, _
    hasfieldnames:=True

DoCmd.TransferSpreadsheet transfertype:=acExport, spreadsheettype:=acSpreadsheetTypeExcel12Xml, _
    TableName:="je", _
    FileName:=strPathSave & strFile, _
    hasfieldnames:=True

Set appXL = New Excel.Application
Set wbCur = appXL.Workbooks.Open(strPathSave & strFile)
Set wbNew = appXL.Workbooks.Add
wbCur.Activate
wbCur.Sheets("je").Copy before:=wbNew.Sheets(1)
wbNew.Activate
appXL.ActiveSheet.Name = "upload"

For Each wsTemp In appXL.Worksheets
    If wsTemp.Name <> "upload" Then
        appXL.Sheets(wsTemp.Name).Delete
    End If
Next wsTemp

appXL.Rows(1).Delete
wbNew.SaveAs "\\tndcfile05\Departments\Accounting\Accounting\JE_Uploads\" & "PO and Toner Reclass.csv", xlCSV
wbNew.Close False

wbCur.Save
wbCur.Close False
appXL.Quit
Set appXL = Nothing

Set db = Nothing
MsgBox "Complete", vbInformation + vbSystemModal, "Complete"

End Sub
