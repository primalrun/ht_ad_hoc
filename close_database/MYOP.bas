Attribute VB_Name = "MYOP"
Option Compare Database

Public Sub InternalSalesMYOP()

Dim intRowLast As Integer
Dim intRowTemp As Integer
Dim intRow As Integer
Dim dblTax As Double
Dim dteStart As Date
Dim dteEnd As Date
Dim strFile As String
Dim wbCur As Excel.Workbook
Dim wbNew As Excel.Workbook
Dim wsJenniferPrintout As Excel.Worksheet
Dim qdefDAO As DAO.QueryDef
Dim rsDAO As DAO.Recordset
Dim fsoTemp As FileSystemObject

dteStart = DLookup("[value_date]", "[global]", "category='close_date_start'")
dteEnd = DLookup("[value_date]", "[global]", "category='close_date_end'")

On Error Resume Next
DoCmd.DeleteObject acTable, "myoptemp"
On Error GoTo 0

strsql = ""
strsql = strsql & vbCrLf & "USE [Playground]" & vbCrLf & _
"" & vbCrLf & _
"DECLARE @RC int" & vbCrLf & _
"DECLARE @datestart date" & vbCrLf & _
"DECLARE @dateend date" & vbCrLf & _
"" & vbCrLf & _
"EXECUTE @RC = [myop\jason.walker].[internal_sales_myop] " & vbCrLf & _
"   '" & CStr(dteStart) & "'" & vbCrLf & _
"  ,'" & CStr(dteEnd) & "'"

Set qdefDAO = CurrentDb.QueryDefs("q_htdw_update")
qdefDAO.SQL = strsql
Set qdefDAO = Nothing

DoCmd.OpenQuery "q_htdw_update"

Set qdefDAO = CurrentDb.QueryDefs("q_htdw_select")
qdefDAO.SQL = "select * from playground.[myop\jason.walker].myoptemp"
Set qdefDAO = Nothing

On Error Resume Next
DoCmd.DeleteObject acTable, "Jennifer_Printout_MYOP"
On Error GoTo 0

On Error Resume Next
DoCmd.DeleteObject acTable, "internal_sales_myop"
On Error GoTo 0


strsql = ""
strsql = strsql & vbCrLf & "select" & vbCrLf & _
"  q.custnbr," & vbCrLf & _
"  q.custname," & vbCrLf & _
"  nz(c.je_category,'MyPoints') as category," & vbCrLf & _
"  round(sum(q.salesnofrt),2) as sales," & vbCrLf & _
"  round(sum(q.loadedcost),2) as cost" & vbCrLf & _
"into Jennifer_Printout_MYOP" & vbCrLf & _
"from q_htdw_select q" & vbCrLf & _
"left join myop_cust_name_je_category c" & vbCrLf & _
"  on q.custname = c.cust_name" & vbCrLf & _
"where" & vbCrLf & _
"  q.primarysalespersoncode not in ('AAAAA', 'BBBBB')" & vbCrLf & _
"group by" & vbCrLf & _
"  q.custnbr," & vbCrLf & _
"  q.custname," & vbCrLf & _
"  nz(c.je_category,'MyPoints')" & vbCrLf & _
"having" & vbCrLf & _
"  (abs(sum(q.salesnofrt))>0 or abs(sum(q.loadedcost))>0)" & vbCrLf & _
"order by 3, 2"

CurrentDb.Execute strsql

strsql = ""
strsql = strsql & vbCrLf & "select" & vbCrLf & _
"  q.category," & vbCrLf & _
"  q.primarysalespersoncode," & vbCrLf & _
"  q.custnbr," & vbCrLf & _
"  q.custname," & vbCrLf & _
"  q.routecd," & vbCrLf & _
"  q.fullinvoicenbr," & vbCrLf & _
"  q.businessunit," & vbCrLf & _
"  q.vendorname," & vbCrLf & _
"  null as loc," & vbCrLf & _
"  null as je_category," & vbCrLf & _
"  round(q.salesnofrt,2) as salesnofrt," & vbCrLf & _
"  round(q.loadedcost,2) as loadedcost" & vbCrLf & _
"into internal_sales_myop" & vbCrLf & _
"from q_htdw_select q"

CurrentDb.Execute strsql

strsql = ""
strsql = strsql & vbCrLf & "update internal_sales_myop ism left join myop_route_code_loc rcl" & vbCrLf & _
"  on ism.routecd = rcl.route_code" & vbCrLf & _
"set ism.loc = rcl.loc" & vbCrLf & _
"where" & vbCrLf & _
"  rcl.route_code is not null"

CurrentDb.Execute strsql

strsql = ""
strsql = strsql & vbCrLf & "update internal_sales_myop ism left join myop_vendor_loc vl" & vbCrLf & _
"  on ism.vendorname = vl.vendor_name" & vbCrLf & _
"set ism.loc = nz(vl.loc,'5003')" & vbCrLf & _
"where" & vbCrLf & _
"  ism.loc is null"

CurrentDb.Execute strsql

strsql = ""
strsql = strsql & vbCrLf & "update internal_sales_myop ism left outer join myop_cust_name_je_category cnc" & vbCrLf & _
"  on ism.custname = cnc.cust_name" & vbCrLf & _
"set ism.je_category = nz(cnc.je_category, 'MyPoints')"

CurrentDb.Execute strsql

'Update category for Seller Flex
CurrentDb.Execute "update internal_sales_myop set category = 'Intercompany' where je_category = 'Seller Flex'"


On Error Resume Next
DoCmd.DeleteObject acTable, "je"
On Error GoTo 0

strsql = ""
strsql = strsql & vbCrLf & "select * into je from (" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  'SJE21' as Doc_Nbr," & vbCrLf & _
"  40000 as GL_Acct," & vbCrLf & _
"  #" & dteEnd & "# as Post_Date," & vbCrLf & _
"  'Write-off I/C Sales' as Description," & vbCrLf & _
"  null as Space_1," & vbCrLf & _
"  'OFFPROD' as BU," & vbCrLf & _
"  cint(i.loc) as Loc," & vbCrLf & _
"  100 as Dept," & vbCrLf & _
"  'MYOP' as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4," & vbCrLf & _
"  round(sum(i.salesnofrt),2) as Amount" & vbCrLf & _
"from internal_sales_myop i" & vbCrLf & _
"where" & vbCrLf & _
"  i.je_category = 'Intercompany'" & vbCrLf & _
"group by"
strsql = strsql & vbCrLf & "  'SJE21'," & vbCrLf & _
"  40000," & vbCrLf & _
"  #" & dteEnd & "#," & vbCrLf & _
"  'Write-off I/C Sales'," & vbCrLf & _
"  null," & vbCrLf & _
"  'OFFPROD'," & vbCrLf & _
"  cint(i.loc)," & vbCrLf & _
"  100," & vbCrLf & _
"  'MYOP'," & vbCrLf & _
"  null," & vbCrLf & _
"  null," & vbCrLf & _
"  null" & vbCrLf & _
"" & vbCrLf & _
"union all" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  'SJE21' as Doc_Nbr," & vbCrLf & _
"  50000 as GL_Acct," & vbCrLf & _
"  #" & dteEnd & "# as Post_Date," & vbCrLf & _
"  'Write-off I/C COGS' as Description,"
strsql = strsql & vbCrLf & "  null as Space_1," & vbCrLf & _
"  'OFFPROD' as BU," & vbCrLf & _
"  cint(i.loc) as Loc," & vbCrLf & _
"  100 as Dept," & vbCrLf & _
"  'MYOP' as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4," & vbCrLf & _
"  -round(sum(i.loadedcost),2) as Amount" & vbCrLf & _
"from internal_sales_myop i" & vbCrLf & _
"group by" & vbCrLf & _
"  'SJE21'," & vbCrLf & _
"  50000," & vbCrLf & _
"  #" & dteEnd & "#," & vbCrLf & _
"  'Write-off I/C COGS'," & vbCrLf & _
"  null," & vbCrLf & _
"  'OFFPROD'," & vbCrLf & _
"  cint(i.loc),"
strsql = strsql & vbCrLf & "  100," & vbCrLf & _
"  'MYOP'," & vbCrLf & _
"  null," & vbCrLf & _
"  null," & vbCrLf & _
"  null" & vbCrLf & _
"" & vbCrLf & _
"union all" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  'SJE21' as Doc_Nbr," & vbCrLf & _
"  65100 as GL_Acct," & vbCrLf & _
"  #" & dteEnd & "# as Post_Date," & vbCrLf & _
"  'MRO Supplies' as Description," & vbCrLf & _
"  null as Space_1," & vbCrLf & _
"  'OFFPROD' as BU," & vbCrLf & _
"  cint(i.loc) as Loc," & vbCrLf & _
"  100 as Dept," & vbCrLf & _
"  'MYOP' as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3,"
strsql = strsql & vbCrLf & "  null as Space_4," & vbCrLf & _
"  round(sum(i.loadedcost),2) as Amount" & vbCrLf & _
"from internal_sales_myop i" & vbCrLf & _
"where" & vbCrLf & _
"  i.je_category = 'MRO'" & vbCrLf & _
"group by" & vbCrLf & _
"  'SJE21'," & vbCrLf & _
"  65100," & vbCrLf & _
"  #" & dteEnd & "#," & vbCrLf & _
"  'MRO Supplies'," & vbCrLf & _
"  null," & vbCrLf & _
"  'OFFPROD'," & vbCrLf & _
"  cint(i.loc)," & vbCrLf & _
"  100," & vbCrLf & _
"  'MYOP'," & vbCrLf & _
"  null," & vbCrLf & _
"  null," & vbCrLf & _
"  null" & vbCrLf & _
"" & vbCrLf & _
"union all"
strsql = strsql & vbCrLf & "" & vbCrLf & _
"select" & vbCrLf & _
"  'SJE21' as Doc_Nbr," & vbCrLf & _
"  21020 as GL_Acct," & vbCrLf & _
"  #" & dteEnd & "# as Post_Date," & vbCrLf & _
"  'MyPoints Expense' as Description," & vbCrLf & _
"  null as Space_1," & vbCrLf & _
"  'OFFPROD' as BU," & vbCrLf & _
"  cint(i.loc) as Loc," & vbCrLf & _
"  100 as Dept," & vbCrLf & _
"  'MYOP' as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4," & vbCrLf & _
"  round(sum(i.loadedcost),2) as Amount" & vbCrLf & _
"from internal_sales_myop i" & vbCrLf & _
"where" & vbCrLf & _
"  i.je_category = 'MyPoints'" & vbCrLf & _
"group by" & vbCrLf & _
"  'SJE21',"
strsql = strsql & vbCrLf & "  21020," & vbCrLf & _
"  #" & dteEnd & "#," & vbCrLf & _
"  'MyPoints Expense'," & vbCrLf & _
"  null," & vbCrLf & _
"  'OFFPROD'," & vbCrLf & _
"  cint(i.loc)," & vbCrLf & _
"  100," & vbCrLf & _
"  'MYOP'," & vbCrLf & _
"  null," & vbCrLf & _
"  null," & vbCrLf & _
"  null" & vbCrLf & _
"" & vbCrLf & _
"union all" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  'SJE21' as Doc_Nbr," & vbCrLf & _
"  61300 as GL_Acct," & vbCrLf & _
"  #" & dteEnd & "# as Post_Date," & vbCrLf & _
"  'Sales Samples' as Description," & vbCrLf & _
"  null as Space_1,"
strsql = strsql & vbCrLf & "  'OFFPROD' as BU," & vbCrLf & _
"  cint(i.loc) as Loc," & vbCrLf & _
"  100 as Dept," & vbCrLf & _
"  'MYOP' as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4," & vbCrLf & _
"  round(sum(i.loadedcost),2) as Amount" & vbCrLf & _
"from internal_sales_myop i" & vbCrLf & _
"where" & vbCrLf & _
"  i.je_category = 'Sales/Marketing'" & vbCrLf & _
"group by" & vbCrLf & _
"  'SJE21'," & vbCrLf & _
"  61300," & vbCrLf & _
"  #" & dteEnd & "#," & vbCrLf & _
"  'Sales Samples'," & vbCrLf & _
"  null," & vbCrLf & _
"  'OFFPROD'," & vbCrLf & _
"  cint(i.loc)," & vbCrLf & _
"  100,"
strsql = strsql & vbCrLf & "  'MYOP'," & vbCrLf & _
"  null," & vbCrLf & _
"  null," & vbCrLf & _
"  null" & vbCrLf & _
"" & vbCrLf & _
"union all" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  'SJE21' as Doc_Nbr," & vbCrLf & _
"  19200 as GL_Acct," & vbCrLf & _
"  #" & dteEnd & "# as Post_Date," & vbCrLf & _
"  'Random Source COGS' as Description," & vbCrLf & _
"  null as Space_1," & vbCrLf & _
"  'OFFPROD' as BU," & vbCrLf & _
"  cint(i.loc) as Loc," & vbCrLf & _
"  100 as Dept," & vbCrLf & _
"  'MYOP' as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4,"
strsql = strsql & vbCrLf & "  round(sum(i.loadedcost),2) as Amount" & vbCrLf & _
"from internal_sales_myop i" & vbCrLf & _
"where" & vbCrLf & _
"  i.je_category = 'Random Source'" & vbCrLf & _
"group by" & vbCrLf & _
"  'SJE21'," & vbCrLf & _
"  19200," & vbCrLf & _
"  #" & dteEnd & "#," & vbCrLf & _
"  'Random Source COGS'," & vbCrLf & _
"  null," & vbCrLf & _
"  'OFFPROD'," & vbCrLf & _
"  cint(i.loc)," & vbCrLf & _
"  100," & vbCrLf & _
"  'MYOP'," & vbCrLf & _
"  null," & vbCrLf & _
"  null," & vbCrLf & _
"  null"


strsql = strsql & vbCrLf & "union all" & vbCrLf & _
"select" & vbCrLf & _
"  'SJE21' as Doc_Nbr," & vbCrLf & _
"  19200 as GL_Acct," & vbCrLf & _
"  #" & dteEnd & "# as Post_Date," & vbCrLf & _
"  'Seller Flex COGS' as Description," & vbCrLf & _
"  null as Space_1," & vbCrLf & _
"  'OFFPROD' as BU," & vbCrLf & _
"  cint(i.loc) as Loc," & vbCrLf & _
"  100 as Dept," & vbCrLf & _
"  'MYOP' as Sell_As," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4,"
strsql = strsql & vbCrLf & "  round(sum(i.loadedcost),2) as Amount" & vbCrLf & _
"from internal_sales_myop i" & vbCrLf & _
"where" & vbCrLf & _
"  i.je_category = 'Seller Flex'" & vbCrLf & _
"group by" & vbCrLf & _
"  'SJE21'," & vbCrLf & _
"  19200," & vbCrLf & _
"  #" & dteEnd & "#," & vbCrLf & _
"  'Seller Flex COGS'," & vbCrLf & _
"  null," & vbCrLf & _
"  'OFFPROD'," & vbCrLf & _
"  cint(i.loc)," & vbCrLf & _
"  100," & vbCrLf & _
"  'MYOP'," & vbCrLf & _
"  null," & vbCrLf & _
"  null," & vbCrLf & _
"  null" & vbCrLf & _
")"


CurrentDb.Execute strsql

CurrentDb.Execute "delete from je where amount = 0"

strsql = ""
strsql = strsql & vbCrLf & "update je j inner join myop_loc_sell_as sa" & vbCrLf & _
"  on j.loc = sa.loc" & vbCrLf & _
"set j.sell_as = sa.sell_as"

CurrentDb.Execute strsql

strsql = "select sum(amount) from je"
Set rsDAO = CurrentDb.OpenRecordset(strsql, dbOpenDynaset)

If rsDAO.BOF = False And rsDAO.EOF = False Then
    rsDAO.MoveFirst
    dblTax = -rsDAO.Fields(0).Value
    
    CurrentDb.Execute "insert into je (Doc_Nbr, GL_Acct, Post_Date, Description, Amount) values('SJE21', 22000, #" & dteEnd & "#, 'Sales Tax Payable', " & round(dblTax, 2) & ")"
End If

rsDAO.Close
Set rsDAO = Nothing

Call PublicStuff.SelectSaveDirectory("Select Directory to Save MYOP Internal Sales Data")

If strPathSave = "" Then
    MsgBox "No Save Directory Selected, Process Cancelled", vbCritical + vbSystemModal, "Save Directory Validation"
    Exit Sub
End If

strFile = "Internal Sales MYOP " & Format(dteStart, "YYYYMM") & ".xlsx"

Set fsoTemp = New FileSystemObject

If fsoTemp.FileExists(strPathSave & strFile) = True Then
    fsoTemp.DeleteFile strPathSave & strFile, True
End If

Set fsoTemp = Nothing

DoCmd.TransferSpreadsheet transfertype:=acExport, _
    spreadsheettype:=acSpreadsheetTypeExcel12Xml, _
    TableName:="Jennifer_Printout_MYOP", _
    FileName:=strPathSave & strFile, _
    hasfieldnames:=True

DoCmd.TransferSpreadsheet transfertype:=acExport, _
    spreadsheettype:=acSpreadsheetTypeExcel12Xml, _
    TableName:="internal_sales_myop", _
    FileName:=strPathSave & strFile, _
    hasfieldnames:=True

DoCmd.TransferSpreadsheet transfertype:=acExport, _
    spreadsheettype:=acSpreadsheetTypeExcel12Xml, _
    TableName:="je", _
    FileName:=strPathSave & strFile, _
    hasfieldnames:=True

'Excel File
'Jennifer Printout
Set appXL = New Excel.Application
Set wbCur = appXL.Workbooks.Open(strPathSave & strFile)
Set wsJenniferPrintout = wbCur.Sheets("Jennifer_Printout_MYOP")
wsJenniferPrintout.Activate
appXL.Rows(1).Delete
intRowLast = appXL.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
intRowTemp = intRowLast

Do Until intRowTemp = 1
    intRow = intRowTemp - 1
    If appXL.Cells(intRowTemp, 3).Value <> appXL.Cells(intRow, 3).Value Then
        appXL.Cells(intRowTemp, 3).EntireRow.Insert
        appXL.Cells(intRowTemp, 3).EntireRow.Insert
    End If
    
    intRowTemp = intRowTemp - 1
Loop

intRowLast = appXL.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
intRowTemp = intRowLast + 1

Do Until intRowTemp = 1
    intRow = intRowTemp - 1
    
    If appXL.Cells(intRow, 3).Value <> "" And appXL.Cells(intRowTemp, 3).Value = "" Then
        appXL.Cells(intRowTemp, 1).Value = appXL.Cells(intRow, 3).Value & " Total"
        appXL.Cells(intRowTemp, 1).Font.Bold = True
        intRows = appXL.Cells(intRow, 3).CurrentRegion.Rows.Count - 1
        appXL.Cells(intRowTemp, 4).FormulaR1C1 = "=sum(r[-" & intRows & "]c:r[-1]c)"
        appXL.Cells(intRowTemp, 5).FormulaR1C1 = "=sum(r[-" & intRows & "]c:r[-1]c)"
        appXL.Range(appXL.Cells(intRowTemp, 4), appXL.Cells(intRowTemp, 5)).Borders(xlEdgeTop).Weight = xlThin
    End If
    
    intRowTemp = intRowTemp - 1
Loop



appXL.Cells(1, 3).EntireColumn.Delete
intRowLast = appXL.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
appXL.Range(appXL.Cells(1, 3), appXL.Cells(intRowLast, 4)).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"

appXL.Cells(1, 1).EntireRow.Insert
appXL.Cells(1, 1).EntireRow.Insert
appXL.Cells(1, 1).EntireRow.Insert

appXL.Cells(1, 1).Value = Format(dteStart, "MMMM YYYY")
appXL.Cells(1, 1).Font.Bold = True
appXL.Range(appXL.Cells(1, 1), appXL.Cells(1, 4)).HorizontalAlignment = xlCenterAcrossSelection

With appXL.Cells(3, 1)
    .Value = "Cust_Nbr"
    .Font.Bold = True
    .HorizontalAlignment = xlLeft
End With

With appXL.Cells(3, 2)
    .Value = "Cust_Name"
    .Font.Bold = True
    .HorizontalAlignment = xlLeft
End With

With appXL.Cells(3, 3)
    .Value = "Sales"
    .Font.Bold = True
    .HorizontalAlignment = xlRight
End With

With appXL.Cells(3, 4)
    .Value = "Cost"
    .Font.Bold = True
    .HorizontalAlignment = xlRight
End With

intRowLast = appXL.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
intRow = intRowLast + 2

appXL.Cells(intRow, 3).FormulaR1C1 = "=sumif(r4c1:r[-1]c1," & Chr(34) & "*Total*" & Chr(34) & ",r4c:r[-1]c)"
appXL.Cells(intRow, 4).FormulaR1C1 = "=sumif(r4c1:r[-1]c1," & Chr(34) & "*Total*" & Chr(34) & ",r4c:r[-1]c)"
appXL.Cells(intRow, 1).Value = "Total"
appXL.Cells(intRow, 1).Font.Bold = True
appXL.Range(appXL.Cells(intRow, 3), appXL.Cells(intRow, 4)).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"

appXL.Columns(1).AutoFit
appXL.Columns(2).AutoFit
appXL.Range(appXL.Cells(1, 3), appXL.Cells(1, 4)).EntireColumn.ColumnWidth = 13
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

If fsoTemp.FileExists("\\tndcfile05\Departments\Accounting\Accounting\JE_Uploads\" & "Internal Sales MYOP.csv") = True Then
    fsoTemp.DeleteFile "\\tndcfile05\Departments\Accounting\Accounting\JE_Uploads\" & "Internal Sales MYOP.csv", True
End If

Set fsoTemp = Nothing

appXL.Rows(1).Delete
wbNew.SaveAs "\\tndcfile05\Departments\Accounting\Accounting\JE_Uploads\" & "Internal Sales MYOP.csv", xlCSV
wbNew.Close False

wbCur.Save
wbCur.Close False
appXL.Quit
Set appXL = Nothing


MsgBox "Success: MYOP Internal Sales is Complete", vbInformation + vbSystemModal, "MYOP Internal Sales"

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

Public Sub SalesAnalysisMYOP()

Dim dteStart As Date
Dim dteEnd As Date
Dim qdefDAO As DAO.QueryDef

dteStart = DLookup("[value_date]", "[global]", "category='close_date_start'")
dteEnd = DLookup("[value_date]", "[global]", "category='close_date_end'")

On Error Resume Next
DoCmd.DeleteObject acTable, "sales_analysis_summary_myop"
On Error GoTo 0

strsql = ""
strsql = strsql & vbCrLf & "SELECT" & vbCrLf & _
"  ad.Company," & vbCrLf & _
"  ad.bu," & vbCrLf & _
"  sum(ad.sales) as Sales," & vbCrLf & _
"  sum(ad.cost) as Cost" & vbCrLf & _
"into sales_analysis_summary_myop" & vbCrLf & _
"from dw_sales_cost_invoice ad" & vbCrLf & _
"WHERE" & vbCrLf & _
"  ad.system = 'Vibe'" & vbCrLf & _
"group BY" & vbCrLf & _
"  ad.Company," & vbCrLf & _
"  ad.bu"

CurrentDb.Execute strsql
Application.RefreshDatabaseWindow
MsgBox "Complete", vbInformation + vbSystemModal, "MYOP Sales Analysis"


End Sub


Public Sub ReclassTSCKittingCOGS()

Dim dteStart As Date
Dim dteEnd As Date
Dim strsql As String
Dim qdefDAO As DAO.QueryDef
Dim db As DAO.Database
Dim fsoTemp As FileSystemObject

dteStart = DLookup("[value_date]", "[global]", "category='close_date_start'")
dteEnd = DLookup("[value_date]", "[global]", "category='close_date_end'")

strsql = ""
strsql = strsql & vbCrLf & "select" & vbCrLf & _
"  sa.cust_name," & vbCrLf & _
"  sum(sa.unloaded_cost) as cost" & vbCrLf & _
"from Playground.[myop\jason.walker].sales_analysis_agg_date sa" & vbCrLf & _
"where" & vbCrLf & _
"  sa.inv_date between '" & CStr(dteStart) & "' and '" & CStr(dteEnd) & "'" & vbCrLf & _
"  and sa.cust_name = 'Tractor Supply - Kitting Program'" & vbCrLf & _
"group by" & vbCrLf & _
"  sa.cust_name"

Set db = CurrentDb

Set fsoTemp = New FileSystemObject

If fsoTemp.FileExists("\\tndcfile05\Departments\Accounting\Accounting\JE_Uploads\" & "Reclass TSC Kitting COGS.csv") = True Then
    fsoTemp.DeleteFile "\\tndcfile05\Departments\Accounting\Accounting\JE_Uploads\" & "Reclass TSC Kitting COGS.csv", True
End If

Set fsoTemp = Nothing

On Error Resume Next
DoCmd.DeleteObject acTable, "je"
DoCmd.DeleteObject acTable, "reclass_tsc_kitting_cogs"
On Error GoTo 0

Set qdefDAO = db.QueryDefs("q_playground_select")
qdefDAO.SQL = strsql
Set qdefDAO = Nothing

db.Execute "select * into reclass_tsc_kitting_cogs from q_playground_select"

strsql = ""
strsql = strsql & vbCrLf & "select * into je from (" & vbCrLf & _
"select" & vbCrLf & _
"  'SJE24' as Doc_Nbr," & vbCrLf & _
"  50000 as GL_Acct," & vbCrLf & _
"  #" & dteEnd & "# as Post_Date," & vbCrLf & _
"  'Reclass TSC Kitting COGS' as Description," & vbCrLf & _
"  null as Space_1," & vbCrLf & _
"  'FULFIL' as BU," & vbCrLf & _
"  '1025' as Loc," & vbCrLf & _
"  '100' as Dept," & vbCrLf & _
"  'MYOP' as Sell_as," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4," & vbCrLf & _
"  -round(p.cost,2) as Amount" & vbCrLf & _
"from q_playground_select p" & vbCrLf & _
"" & vbCrLf & _
"union all" & vbCrLf & _
"" & vbCrLf & _
"select"
strsql = strsql & vbCrLf & "  'SJE24' as Doc_Nbr," & vbCrLf & _
"  51000 as GL_Acct," & vbCrLf & _
"  #" & dteEnd & "# as Post_Date," & vbCrLf & _
"  'Reclass TSC Kitting COGS' as Description," & vbCrLf & _
"  null as Space_1," & vbCrLf & _
"  'FULFIL' as BU," & vbCrLf & _
"  '1025' as Loc," & vbCrLf & _
"  '100' as Dept," & vbCrLf & _
"  'MYOP' as Sell_as," & vbCrLf & _
"  null as Space_2," & vbCrLf & _
"  null as Space_3," & vbCrLf & _
"  null as Space_4," & vbCrLf & _
"  round(p.cost,2) as Amount" & vbCrLf & _
"from q_playground_select p) q1"

db.Execute strsql
Application.RefreshDatabaseWindow

Call PublicStuff.SelectSaveDirectory("Select Directory to Save Reclass TSC Kitting COGS")

If strPathSave = "" Then
    MsgBox "No Save Directory Selected, Process Cancelled", vbCritical + vbSystemModal, "Save Directory Validation"
    Exit Sub
End If

strFile = "Reclass TSC Kitting COGS " & Format(dteStart, "YYYYMM") & ".xlsx"

Set fsoTemp = New FileSystemObject

If fsoTemp.FileExists(strPathSave & strFile) = True Then
    fsoTemp.DeleteFile strPathSave & strFile, True
End If

Set fsoTemp = Nothing

DoCmd.TransferSpreadsheet transfertype:=acExport, _
    spreadsheettype:=acSpreadsheetTypeExcel12Xml, _
    TableName:="reclass_tsc_kitting_cogs", _
    FileName:=strPathSave & strFile, _
    hasfieldnames:=True

DoCmd.TransferSpreadsheet transfertype:=acExport, _
    spreadsheettype:=acSpreadsheetTypeExcel12Xml, _
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
wbNew.SaveAs "\\tndcfile05\Departments\Accounting\Accounting\JE_Uploads\" & "Reclass TSC Kitting COGS.csv", xlCSV
wbNew.Close False

wbCur.Save
wbCur.Close False
appXL.Quit
Set appXL = Nothing


Set db = Nothing
MsgBox "Complete", vbInformation + vbSystemModal, "Complete"

End Sub
