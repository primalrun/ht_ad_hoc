Attribute VB_Name = "LineVariances"
Option Compare Database
Option Explicit

Public Sub HTBuild()

Dim dteStart As Date
Dim dteEnd As Date
Dim strsql As String
Dim db As DAO.Database

Set db = CurrentDb
db.Execute "delete from line_variance_ht"
dteStart = DLookup("[value_date]", "[global]", "category='close_date_start'")
dteEnd = DLookup("[value_date]", "[global]", "category='close_date_end'")

'MYOI Prior Month Accrual
strsql = ""
strsql = strsql & vbCrLf & "insert into line_variance_ht (description, sales_total, cost_total, sales_myoi, cost_myoi)" & vbCrLf & _
"select" & vbCrLf & _
"  'MYOI Prior Month Accrual'," & vbCrLf & _
"  -e.sales," & vbCrLf & _
"  -e.cost," & vbCrLf & _
"  -e.sales," & vbCrLf & _
"  -e.cost" & vbCrLf & _
"from ecommerce_accrual_trxn e" & vbCrLf & _
"where" & vbCrLf & _
"  e.post_date = #" & DateAdd("d", -1, dteStart) & "#"

db.Execute strsql

'Plus I/C Sales & IPRs Due from RAC
strsql = ""
strsql = strsql & vbCrLf & "insert into line_variance_ht (description, sales_total, cost_total, sales_op, cost_op)" & vbCrLf & _
"select" & vbCrLf & _
"  'Plus I/C Sales & IPRs Due from RAC'," & vbCrLf & _
"  sum(iif(v.type = 'Invoice' and  v.category = 'MRO - RAC' and v.measure = 'Sales' and v.org = 'HT', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Invoice' and  v.category = 'MRO - RAC' and v.measure = 'Cost' and v.org = 'HT', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Invoice' and  v.category = 'MRO - RAC' and v.measure = 'Sales' and v.org = 'HT', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Invoice' and  v.category = 'MRO - RAC' and v.measure = 'Cost' and v.org = 'HT', v.variance, null))" & vbCrLf & _
"from variance_all v"

db.Execute strsql

'Plus I/C Sales & IPRs Intercompany
strsql = ""
strsql = strsql & vbCrLf & "insert into line_variance_ht (description, sales_total, cost_total, sales_op, cost_op)" & vbCrLf & _
"select" & vbCrLf & _
"  'Plus I/C Sales & IPRs Intercompany'," & vbCrLf & _
"  sum(iif(v.type = 'Invoice' and  v.category = 'Intercompany' and v.measure = 'Sales' and v.org = 'HT', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Invoice' and  v.category = 'Intercompany' and v.measure = 'Cost' and v.org = 'HT', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Invoice' and  v.category = 'Intercompany' and v.measure = 'Sales' and v.org = 'HT', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Invoice' and  v.category = 'Intercompany' and v.measure = 'Cost' and v.org = 'HT', v.variance, null))" & vbCrLf & _
"from variance_all v"

db.Execute strsql

'Plus I/C Sales & IPRs MRO - Corp IT
strsql = ""
strsql = strsql & vbCrLf & "insert into line_variance_ht (description, sales_total, cost_total, sales_op, cost_op)" & vbCrLf & _
"select" & vbCrLf & _
"  'Plus I/C Sales & IPRs MRO - Corp IT'," & vbCrLf & _
"  sum(iif(v.type = 'Invoice' and  v.category = 'MRO - Corp IT' and v.measure = 'Sales' and v.org = 'HT', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Invoice' and  v.category = 'MRO - Corp IT' and v.measure = 'Cost' and v.org = 'HT', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Invoice' and  v.category = 'MRO - Corp IT' and v.measure = 'Sales' and v.org = 'HT', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Invoice' and  v.category = 'MRO - Corp IT' and v.measure = 'Cost' and v.org = 'HT', v.variance, null))" & vbCrLf & _
"from variance_all v"

db.Execute strsql

'Plus I/C Sales & IPRs MRO - HT
strsql = ""
strsql = strsql & vbCrLf & "insert into line_variance_ht (description, sales_total, cost_total, sales_op, cost_op)" & vbCrLf & _
"select" & vbCrLf & _
"  'Plus I/C Sales & IPRs MRO - HT'," & vbCrLf & _
"  sum(iif(v.type = 'Invoice' and  v.category = 'MRO - HT' and v.measure = 'Sales' and v.org = 'HT', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Invoice' and  v.category = 'MRO - HT' and v.measure = 'Cost' and v.org = 'HT', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Invoice' and  v.category = 'MRO - HT' and v.measure = 'Sales' and v.org = 'HT', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Invoice' and  v.category = 'MRO - HT' and v.measure = 'Cost' and v.org = 'HT', v.variance, null))" & vbCrLf & _
"from variance_all v"

db.Execute strsql

'Plus I/C Sales & IPRs Sales/Marketing
strsql = ""
strsql = strsql & vbCrLf & "insert into line_variance_ht (description, sales_total, cost_total, sales_op, cost_op)" & vbCrLf & _
"select" & vbCrLf & _
"  'Plus I/C Sales & IPRs Sales/Marketing'," & vbCrLf & _
"  sum(iif(v.type = 'Invoice' and  v.category = 'Sales/Marketing' and v.measure = 'Sales' and v.org = 'HT', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Invoice' and  v.category = 'Sales/Marketing' and v.measure = 'Cost' and v.org = 'HT', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Invoice' and  v.category = 'Sales/Marketing' and v.measure = 'Sales' and v.org = 'HT', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Invoice' and  v.category = 'Sales/Marketing' and v.measure = 'Cost' and v.org = 'HT', v.variance, null))" & vbCrLf & _
"from variance_all v"

db.Execute strsql

'Credit Memos
strsql = ""
strsql = strsql & vbCrLf & "insert into line_variance_ht (description, sales_total, cost_total, sales_dw, cost_dw, sales_ff, cost_ff, sales_its, cost_its, sales_mps, cost_mps, sales_myoi, cost_myoi, sales_op, cost_op)" & vbCrLf & _
"select" & vbCrLf & _
"  'Credit Memos'," & vbCrLf & _
"  sum(iif(v.type in ('Credit', 'Refund') and v.measure = 'Sales' and v.org = 'HT', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type in ('Credit', 'Refund') and v.measure = 'Cost' and v.org = 'HT', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type in ('Credit', 'Refund') and v.measure = 'Sales' and v.org = 'HT' and v.bu = 'DWORKS', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type in ('Credit', 'Refund') and v.measure = 'Cost' and v.org = 'HT' and v.bu = 'DWORKS', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type in ('Credit', 'Refund') and v.measure = 'Sales' and v.org = 'HT' and v.bu = 'FULFIL', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type in ('Credit', 'Refund') and v.measure = 'Cost' and v.org = 'HT' and v.bu = 'FULFIL', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type in ('Credit', 'Refund') and v.measure = 'Sales' and v.org = 'HT' and v.bu = 'ITS', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type in ('Credit', 'Refund') and v.measure = 'Cost' and v.org = 'HT' and v.bu = 'ITS', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type in ('Credit', 'Refund') and v.measure = 'Sales' and v.org = 'HT' and v.bu = 'MPS', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type in ('Credit', 'Refund') and v.measure = 'Cost' and v.org = 'HT' and v.bu = 'MPS', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type in ('Credit', 'Refund') and v.measure = 'Sales' and v.org = 'HT' and v.bu = 'MYOI', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type in ('Credit', 'Refund') and v.measure = 'Cost' and v.org = 'HT' and v.bu = 'MYOI', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type in ('Credit', 'Refund') and v.measure = 'Sales' and v.org = 'HT' and v.bu = 'OFFPROD', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type in ('Credit', 'Refund') and v.measure = 'Cost' and v.org = 'HT' and v.bu = 'OFFPROD', v.variance, null))" & vbCrLf & _
"from variance_all v"

db.Execute strsql

'Return
strsql = ""
strsql = strsql & vbCrLf & "insert into line_variance_ht (description, sales_total, cost_total, sales_dw, cost_dw, sales_ff, cost_ff, sales_its, cost_its, sales_mps, cost_mps, sales_myoi, cost_myoi, sales_op, cost_op)" & vbCrLf & _
"select" & vbCrLf & _
"  'Returns'," & vbCrLf & _
"  sum(iif(v.type = 'Return' and v.measure = 'Sales' and v.org = 'HT', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Return' and v.measure = 'Cost' and v.org = 'HT', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Return' and v.measure = 'Sales' and v.org = 'HT' and v.bu = 'DWORKS', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Return' and v.measure = 'Cost' and v.org = 'HT' and v.bu = 'DWORKS', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Return' and v.measure = 'Sales' and v.org = 'HT' and v.bu = 'FULFIL', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Return' and v.measure = 'Cost' and v.org = 'HT' and v.bu = 'FULFIL', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Return' and v.measure = 'Sales' and v.org = 'HT' and v.bu = 'ITS', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Return' and v.measure = 'Cost' and v.org = 'HT' and v.bu = 'ITS', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Return' and v.measure = 'Sales' and v.org = 'HT' and v.bu = 'MPS', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Return' and v.measure = 'Cost' and v.org = 'HT' and v.bu = 'MPS', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Return' and v.measure = 'Sales' and v.org = 'HT' and v.bu = 'MYOI', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Return' and v.measure = 'Cost' and v.org = 'HT' and v.bu = 'MYOI', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Return' and v.measure = 'Sales' and v.org = 'HT' and v.bu = 'OFFPROD', v.variance, null))," & vbCrLf & _
"  sum(iif(v.type = 'Return' and v.measure = 'Cost' and v.org = 'HT' and v.bu = 'OFFPROD', v.variance, null))" & vbCrLf & _
"from variance_all v"

db.Execute strsql


'MPS-Montefiore Copier Program
strsql = ""
strsql = strsql & vbCrLf & "insert into line_variance_ht (description, sales_total, cost_total, sales_mps, cost_mps)" & vbCrLf & _
"select" & vbCrLf & _
"  'MPS-Montefiore Copier Program',  " & vbCrLf & _
"  -sum(s.cost)," & vbCrLf & _
"  -sum(s.cost)," & vbCrLf & _
"  -sum(s.cost)," & vbCrLf & _
"  -sum(s.cost)" & vbCrLf & _
"from dw_sales_cost_invoice s" & vbCrLf & _
"where" & vbCrLf & _
"  s.custname = 'MPS-Montefiore Copier Program'"

db.Execute strsql


'MPS-Montefiore Mt Vernon Copiers
strsql = ""
strsql = strsql & vbCrLf & "insert into line_variance_ht (description, sales_total, cost_total, sales_mps, cost_mps)" & vbCrLf & _
"select" & vbCrLf & _
"  'MPS-Montefiore Mt Vernon Copiers',  " & vbCrLf & _
"  -sum(s.cost)," & vbCrLf & _
"  -sum(s.cost)," & vbCrLf & _
"  -sum(s.cost)," & vbCrLf & _
"  -sum(s.cost)" & vbCrLf & _
"from dw_sales_cost_invoice s" & vbCrLf & _
"where" & vbCrLf & _
"  s.custname = 'MPS-Montefiore Mt Vernon Copiers'"

db.Execute strsql


'MPS-Montefiore New Rochelle Copiers
strsql = ""
strsql = strsql & vbCrLf & "insert into line_variance_ht (description, sales_total, cost_total, sales_mps, cost_mps)" & vbCrLf & _
"select" & vbCrLf & _
"  'MPS-Montefiore Copier Program',  " & vbCrLf & _
"  -sum(s.cost)," & vbCrLf & _
"  -sum(s.cost)," & vbCrLf & _
"  -sum(s.cost)," & vbCrLf & _
"  -sum(s.cost)" & vbCrLf & _
"from dw_sales_cost_invoice s" & vbCrLf & _
"where" & vbCrLf & _
"  s.custname = 'MPS-Montefiore New Rochelle Copiers'"

db.Execute strsql






End Sub
