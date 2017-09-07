Attribute VB_Name = "DW"
Option Compare Database
Option Explicit

Public Sub DWSalesProduct()

Dim dteStart As Date
Dim dteEnd As Date
Dim strsql As String
Dim qdefDAO As DAO.QueryDef
Dim db As DAO.Database

Set db = CurrentDb
dteStart = DLookup("[value_date]", "[global]", "category='close_date_start'")
dteEnd = DLookup("[value_date]", "[global]", "category='close_date_end'")

db.Execute "delete from dw_sales_cost_invoice"

strsql = ""
strsql = strsql & vbCrLf & "USE [Playground]" & vbCrLf & _
"" & vbCrLf & _
"DECLARE @RC int" & vbCrLf & _
"DECLARE @datestart date" & vbCrLf & _
"DECLARE @dateend date" & vbCrLf & _
"" & vbCrLf & _
"EXECUTE @RC = [myop\jason.walker].[sales_analysis_invoice_dw_monthly] " & vbCrLf & _
"   '" & CStr(dteStart) & "'" & vbCrLf & _
"  ,'" & CStr(dteEnd) & "'"

Set qdefDAO = CurrentDb.QueryDefs("q_playground_update")
qdefDAO.SQL = strsql
Set qdefDAO = Nothing

DoCmd.OpenQuery "q_playground_update"


strsql = ""
strsql = strsql & vbCrLf & "select" & vbCrLf & _
"  sam.businessunit as bu," & vbCrLf & _
"  sam.doctypedesc," & vbCrLf & _
"  sam.fullinvoicenbr as invoice," & vbCrLf & _
"  sam.systemname as system," & vbCrLf & _
"  sam.company," & vbCrLf & _
"  sam.custnbr," & vbCrLf & _
"  sam.custname," & vbCrLf & _
"  sum(sam.salesnofrt) as sales," & vbCrLf & _
"  sum(sam.totalunloadedcost) as cost" & vbCrLf & _
"from Playground.[myop\jason.walker].sales_analysis_monthly_dw sam" & vbCrLf & _
"group by" & vbCrLf & _
"  sam.businessunit," & vbCrLf & _
"  sam.doctypedesc," & vbCrLf & _
"  sam.fullinvoicenbr," & vbCrLf & _
"  sam.systemname," & vbCrLf & _
"  sam.company," & vbCrLf & _
"  sam.custnbr," & vbCrLf & _
"  sam.custname" & vbCrLf & _
"having (abs(sum(sam.salesnofrt))>0 or abs(sum(sam.totalunloadedcost))>0)"

Set qdefDAO = CurrentDb.QueryDefs("q_playground_select")
qdefDAO.SQL = strsql
Set qdefDAO = Nothing

On Error Resume Next
DoCmd.DeleteObject acQuery, "t1"
On Error GoTo 0

strsql = ""
strsql = strsql & vbCrLf & "select" & vbCrLf & _
"  q.bu," & vbCrLf & _
"  q.doctypedesc," & vbCrLf & _
"  q.invoice," & vbCrLf & _
"  q.system," & vbCrLf & _
"  q.company," & vbCrLf & _
"  q.custnbr," & vbCrLf & _
"  q.custname," & vbCrLf & _
"  b.bu_gl," & vbCrLf & _
"  cdbl(sum(q.sales)) as sales," & vbCrLf & _
"  cdbl(sum(q.cost)) as cost" & vbCrLf & _
"from q_playground_select q" & vbCrLf & _
"  left join bu_conversion_gl_dw b" & vbCrLf & _
"    on q.bu = b.bu_dw" & vbCrLf & _
"group by" & vbCrLf & _
"  q.bu," & vbCrLf & _
"  q.doctypedesc," & vbCrLf & _
"  q.invoice," & vbCrLf & _
"  q.system,"
strsql = strsql & vbCrLf & "  q.company," & vbCrLf & _
"  q.custnbr," & vbCrLf & _
"  q.custname," & vbCrLf & _
"  b.bu_gl"

Set qdefDAO = db.CreateQueryDef("t1", strsql)
db.Execute "insert into dw_sales_cost_invoice select * from t1"
Set qdefDAO = Nothing
DoCmd.DeleteObject acQuery, "t1"

Application.RefreshDatabaseWindow

CurrentDb.Execute "update global set value_date = now where category = 'dw_sales_product_update'"
MsgBox "Complete: Update DW Sales Product Data", vbInformation + vbSystemModal, "Update DW Sales Product Data"

End Sub

