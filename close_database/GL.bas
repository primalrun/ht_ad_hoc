Attribute VB_Name = "GL"
Option Compare Database

Public Sub UpdateGLSalesandCOGS()

Dim dteStart As Date
Dim dteEnd As Date
Dim strsql As String
Dim qdefDAO As DAO.QueryDef
Dim db As DAO.Database

Set db = CurrentDb
dteStart = DLookup("[value_date]", "[global]", "category='close_date_start'")
dteEnd = DLookup("[value_date]", "[global]", "category='close_date_end'")

strsql = ""
strsql = strsql & vbCrLf & "USE [Playground]" & vbCrLf & _
"" & vbCrLf & _
"DECLARE @RC int" & vbCrLf & _
"" & vbCrLf & _
"EXECUTE @RC = [myop\jason.walker].[proc_gl_account_reporting]"

Set qdefDAO = CurrentDb.QueryDefs("q_playground_update")
qdefDAO.SQL = strsql
Set qdefDAO = Nothing
DoCmd.OpenQuery "q_playground_update"

strsql = "set nocount on"
strsql = strsql & vbCrLf & "declare @datestart date = '" & dteStart & "'" & vbCrLf & _
"declare @dateend date = '" & dteEnd & "'" & vbCrLf & _
"" & vbCrLf & _
"delete from htdw.Playground.[myop\jason.walker].sales_cost_gl" & vbCrLf & _
"" & vbCrLf & _
"INSERT INTO htdw.Playground.[myop\jason.walker].sales_cost_gl" & vbCrLf & _
"SELECT" & vbCrLf & _
"@datestart AS RptPeriod," & vbCrLf & _
"'HT' AS Org," & vbCrLf & _
"G.[Global Dimension 2 Code] as BU," & vbCrLf & _
"G.[Document No_]," & vbCrLf & _
"CASE G.[Document Type]" & vbCrLf & _
"WHEN 1 THEN 'Payment'" & vbCrLf & _
"WHEN 2 THEN 'Invoice'" & vbCrLf & _
"WHEN 3 THEN 'Credit'" & vbCrLf & _
"WHEN 4 THEN 'Finance Charge'" & vbCrLf & _
"WHEN 5 THEN 'Reminder'" & vbCrLf & _
"WHEN 6 THEN 'Refund'" & vbCrLf & _
"ELSE 'Invoice'" & vbCrLf & _
"END AS Type,  "
strsql = strsql & vbCrLf & "SUM(CASE WHEN glr.level_1 = 'Sales' THEN G.Amount ELSE 0 END) * -1 AS Sales_Amount," & vbCrLf & _
"SUM(CASE WHEN glr.level_1 = 'Cost' THEN G.Amount ELSE 0 END) AS Cost_Amount" & vbCrLf & _
"FROM NAVRep.dbo.[Hi Touch$G_L Entry] G WITH(NOLOCK)" & vbCrLf & _
"inner join htdw.Playground.[myop\jason.walker].gl_account_reporting glr" & vbCrLf & _
"on G.[G_L Account No_] = glr.gl_account" & vbCrLf & _
"and glr.level_1 in ('Sales', 'Cost')" & vbCrLf & _
"and glr.company = 'HT'" & vbCrLf & _
"WHERE" & vbCrLf & _
"G.[Posting Date] BETWEEN @datestart AND @dateend  " & vbCrLf & _
"GROUP BY" & vbCrLf & _
"G.[Global Dimension 2 Code],  " & vbCrLf & _
"G.[Document No_]," & vbCrLf & _
"CASE G.[Document Type]" & vbCrLf & _
"WHEN 1 THEN 'Payment'" & vbCrLf & _
"WHEN 2 THEN 'Invoice'" & vbCrLf & _
"WHEN 3 THEN 'Credit'" & vbCrLf & _
"WHEN 4 THEN 'Finance Charge'" & vbCrLf & _
"WHEN 5 THEN 'Reminder'" & vbCrLf & _
"WHEN 6 THEN 'Refund'" & vbCrLf & _
"ELSE 'Invoice'"
strsql = strsql & vbCrLf & "END" & vbCrLf & _
"" & vbCrLf & _
"INSERT INTO htdw.Playground.[myop\jason.walker].sales_cost_gl" & vbCrLf & _
"SELECT" & vbCrLf & _
"@datestart AS RptPeriod," & vbCrLf & _
"'MYOP' AS Org," & vbCrLf & _
"G.[Global Dimension 2 Code] as BU," & vbCrLf & _
"G.[Document No_]," & vbCrLf & _
"CASE G.[Document Type]" & vbCrLf & _
"WHEN 1 THEN 'Payment'" & vbCrLf & _
"WHEN 2 THEN 'Invoice'" & vbCrLf & _
"WHEN 3 THEN 'Credit'" & vbCrLf & _
"WHEN 4 THEN 'Finance Charge'" & vbCrLf & _
"WHEN 5 THEN 'Reminder'" & vbCrLf & _
"WHEN 6 THEN 'Refund'" & vbCrLf & _
"ELSE 'Invoice'" & vbCrLf & _
"END AS Type," & vbCrLf & _
"SUM(CASE WHEN glr.level_1 = 'Sales' and glr.gl_account <> '41000'THEN G.Amount ELSE 0 END) * -1 AS Sales_Amount," & vbCrLf & _
"SUM(CASE WHEN glr.level_1 = 'Cost' THEN G.Amount ELSE 0 END) AS Cost_Amount" & vbCrLf & _
"FROM NAVRep.dbo.[MYOP$G_L Entry] G WITH(NOLOCK)"
strsql = strsql & vbCrLf & "inner join htdw.Playground.[myop\jason.walker].gl_account_reporting glr" & vbCrLf & _
"on G.[G_L Account No_] = glr.gl_account" & vbCrLf & _
"and glr.level_1 in ('Sales', 'Cost')" & vbCrLf & _
"and glr.company = 'MYOP'" & vbCrLf & _
"WHERE" & vbCrLf & _
"G.[Posting Date] BETWEEN @datestart AND @dateend  " & vbCrLf & _
"GROUP BY" & vbCrLf & _
"G.[Global Dimension 2 Code],  " & vbCrLf & _
"G.[Document No_]," & vbCrLf & _
"CASE G.[Document Type]" & vbCrLf & _
"WHEN 1 THEN 'Payment'" & vbCrLf & _
"WHEN 2 THEN 'Invoice'" & vbCrLf & _
"WHEN 3 THEN 'Credit'" & vbCrLf & _
"WHEN 4 THEN 'Finance Charge'" & vbCrLf & _
"WHEN 5 THEN 'Reminder'" & vbCrLf & _
"WHEN 6 THEN 'Refund'" & vbCrLf & _
"ELSE 'Invoice'" & vbCrLf & _
"END"

Set qdefDAO = CurrentDb.QueryDefs("q_navrep_update")
qdefDAO.SQL = strsql
Set qdefDAO = Nothing
DoCmd.OpenQuery "q_navrep_update"

db.Execute "delete from gl_sales_cogs"

Set qdefDAO = db.QueryDefs("q_playground_select")
qdefDAO.SQL = "select *, 'None' as category from playground.[myop\jason.walker].sales_cost_gl"
Set qdefDAO = Nothing

strsql = ""
strsql = strsql & vbCrLf & "insert into gl_sales_cogs "
strsql = strsql & vbCrLf & "select" & vbCrLf & _
"  g.rptperiod," & vbCrLf & _
"  g.org," & vbCrLf & _
"  g.bu," & vbCrLf & _
"  g.invoice," & vbCrLf & _
"  g.doctype," & vbCrLf & _
"  cdbl(sum(g.sales)) as sales," & vbCrLf & _
"  cdbl(sum(g.cost)) as cost," & vbCrLf & _
"  g.category" & vbCrLf & _
"from q_playground_select g" & vbCrLf & _
"group by" & vbCrLf & _
"  g.rptperiod," & vbCrLf & _
"  g.org," & vbCrLf & _
"  g.bu," & vbCrLf & _
"  g.invoice," & vbCrLf & _
"  g.doctype," & vbCrLf & _
"  g.category" & vbCrLf & _
"having (sum(g.sales)<>0 or sum(g.cost) <> 0)"

db.Execute strsql

db.Execute "update gl_sales_cogs set bu = 'OFFPROD' where bu is null or len(bu) = 0"

Set db = Nothing
MsgBox "Success", vbInformation + vbSystemModal, "Complete"

End Sub


Public Sub CheckForDocNumbers()

Dim dteStart As Date
Dim dteEnd As Date
Dim strsql As String
Dim qdefDAO As DAO.QueryDef
Dim db As DAO.Database

Set db = CurrentDb
dteStart = DLookup("[value_date]", "[global]", "category='close_date_start'")
dteEnd = DLookup("[value_date]", "[global]", "category='close_date_end'")

strsql = "set nocount on"
strsql = strsql & vbCrLf & "declare @datestart date = '" & CStr(dteStart) & "'" & vbCrLf & _
"declare @dateend date = '" & CStr(dteEnd) & "'" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  'SuperWarehouse' as 'SA_Description'," & vbCrLf & _
"  gl.[G_L Account No_]," & vbCrLf & _
"  gl.[Posting Date]," & vbCrLf & _
"  gl.[Document No_]," & vbCrLf & _
"  gl.Description," & vbCrLf & _
"  gl.[User ID]," & vbCrLf & _
"  sum(case when gar.level_1 = 'Sales' and gar.company = 'HT' then gl.Amount else 0 end) * -1 AS SalesAmount," & vbCrLf & _
"  -sum(case when gar.level_1 = 'Cost' and gar.company = 'HT' then gl.Amount else 0 end) * -1 AS CostAmount" & vbCrLf & _
"from NAVRep.dbo.[Hi Touch$G_L Entry] gl with(nolock)" & vbCrLf & _
"  inner join HTDW.Playground.[myop\jason.walker].gl_account_reporting gar" & vbCrLf & _
"    on gl.[G_L Account No_] = gar.gl_account" & vbCrLf & _
"where" & vbCrLf & _
"  gl.[Posting Date] between @datestart and @dateend" & vbCrLf & _
"  and gar.level_1 in ('Sales', 'Cost')" & vbCrLf & _
"  and (gl.Description like '%Super%'" & vbCrLf & _
"  or gl.Description like '%SW%')"
strsql = strsql & vbCrLf & "group by" & vbCrLf & _
"  gl.[G_L Account No_]," & vbCrLf & _
"  gl.[Posting Date]," & vbCrLf & _
"  gl.[Document No_]," & vbCrLf & _
"  gl.Description," & vbCrLf & _
"  gl.[User ID]" & vbCrLf & _
"" & vbCrLf & _
"union all" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  'RandomSource' as 'SA_Description'," & vbCrLf & _
"  gl.[G_L Account No_]," & vbCrLf & _
"  gl.[Posting Date]," & vbCrLf & _
"  gl.[Document No_]," & vbCrLf & _
"  gl.Description," & vbCrLf & _
"  gl.[User ID]," & vbCrLf & _
"  sum(case when gar.level_1 = 'Sales' and gar.company = 'HT' then gl.Amount else 0 end) * -1 AS SalesAmount," & vbCrLf & _
"  -sum(case when gar.level_1 = 'Cost' and gar.company = 'HT' then gl.Amount else 0 end) * -1 AS CostAmount" & vbCrLf & _
"from NAVRep.dbo.[Hi Touch$G_L Entry] gl with(nolock)" & vbCrLf & _
"  inner join HTDW.Playground.[myop\jason.walker].gl_account_reporting gar"
strsql = strsql & vbCrLf & "  on gl.[G_L Account No_] = gar.gl_account" & vbCrLf & _
"where" & vbCrLf & _
"  gl.[Posting Date] between @datestart and @dateend" & vbCrLf & _
"  and gl.Description like '%Random%'" & vbCrLf & _
"  and gar.level_1 in ('Sales', 'Cost')" & vbCrLf & _
"group by" & vbCrLf & _
"  gl.[G_L Account No_]," & vbCrLf & _
"  gl.[Posting Date]," & vbCrLf & _
"  gl.[Document No_]," & vbCrLf & _
"  gl.Description," & vbCrLf & _
"  gl.[User ID]" & vbCrLf & _
"" & vbCrLf & _
"union all" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  'Condeco' as 'SA_Description'," & vbCrLf & _
"  gl.[G_L Account No_]," & vbCrLf & _
"  gl.[Posting Date]," & vbCrLf & _
"  gl.[Document No_]," & vbCrLf & _
"  gl.Description,"
strsql = strsql & vbCrLf & "  gl.[User ID]," & vbCrLf & _
"  sum(case when gar.level_1 = 'Sales' and gar.company = 'HT' then gl.Amount else 0 end) * -1 AS SalesAmount," & vbCrLf & _
"  -sum(case when gar.level_1 = 'Cost' and gar.company = 'HT' then gl.Amount else 0 end) * -1 AS CostAmount" & vbCrLf & _
"from NAVRep.dbo.[Hi Touch$G_L Entry] gl with(nolock)" & vbCrLf & _
"  inner join HTDW.Playground.[myop\jason.walker].gl_account_reporting gar" & vbCrLf & _
"  on gl.[G_L Account No_] = gar.gl_account" & vbCrLf & _
"where" & vbCrLf & _
"  gl.[Posting Date] between @datestart and @dateend" & vbCrLf & _
"  and gl.Description like '%condeco%'" & vbCrLf & _
"  and gar.level_1 in ('Sales', 'Cost')" & vbCrLf & _
"group by" & vbCrLf & _
"  gl.[G_L Account No_]," & vbCrLf & _
"  gl.[Posting Date]," & vbCrLf & _
"  gl.[Document No_]," & vbCrLf & _
"  gl.Description," & vbCrLf & _
"  gl.[User ID]" & vbCrLf & _
"" & vbCrLf & _
"  union all" & vbCrLf & _
"" & vbCrLf & _
"  select"
strsql = strsql & vbCrLf & "  'DS Services' as 'SA_Description'," & vbCrLf & _
"  gl.[G_L Account No_]," & vbCrLf & _
"  gl.[Posting Date]," & vbCrLf & _
"  gl.[Document No_]," & vbCrLf & _
"  gl.Description," & vbCrLf & _
"  gl.[User ID]," & vbCrLf & _
"  sum(case when gar.level_1 = 'Sales' and gar.company = 'HT' then gl.Amount else 0 end) * -1 AS SalesAmount," & vbCrLf & _
"  -sum(case when gar.level_1 = 'Cost' and gar.company = 'HT' then gl.Amount else 0 end) * -1 AS CostAmount" & vbCrLf & _
"from NAVRep.dbo.[Hi Touch$G_L Entry] gl with(nolock)" & vbCrLf & _
"  inner join HTDW.Playground.[myop\jason.walker].gl_account_reporting gar" & vbCrLf & _
"  on gl.[G_L Account No_] = gar.gl_account" & vbCrLf & _
"where" & vbCrLf & _
"  gl.[Posting Date] between @datestart and @dateend" & vbCrLf & _
"  and gl.Description like '%ds service%'" & vbCrLf & _
"  and gar.level_1 in ('Sales', 'Cost')" & vbCrLf & _
"group by" & vbCrLf & _
"  gl.[G_L Account No_]," & vbCrLf & _
"  gl.[Posting Date]," & vbCrLf & _
"  gl.[Document No_]," & vbCrLf & _
"  gl.Description,"
strsql = strsql & vbCrLf & "  gl.[User ID]"


Set qdefDAO = db.QueryDefs("q_navrep_select")
qdefDAO.SQL = strsql
Set qdefDAO = Nothing
DoCmd.OpenQuery "q_htdw_select"


Set db = Nothing

End Sub
