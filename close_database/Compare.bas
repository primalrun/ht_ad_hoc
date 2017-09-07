Attribute VB_Name = "Compare"
Option Compare Database

Public Sub Comparison()

Dim strsql As String
Dim db As DAO.Database
Dim qdefDAO As DAO.QueryDef

Set db = CurrentDb
dteStart = DLookup("[value_date]", "[global]", "category='close_date_start'")
dteEnd = DLookup("[value_date]", "[global]", "category='close_date_end'")


On Error Resume Next
DoCmd.DeleteObject acTable, "ht_invoice_internal"
DoCmd.DeleteObject acTable, "myop_invoice_internal"
DoCmd.DeleteObject acQuery, "t1"
On Error GoTo 0

db.Execute "delete from sales_variance_ht"
db.Execute "delete from cost_variance_ht"
db.Execute "delete from sales_variance_myop"
db.Execute "delete from cost_variance_myop"
db.Execute "delete from variance_all"

'HT Invoice Internal
strsql = ""
strsql = strsql & vbCrLf & "select" & vbCrLf & _
"  *" & vbCrLf & _
"into ht_invoice_internal" & vbCrLf & _
"from (" & vbCrLf & _
"select distinct" & vbCrLf & _
"  i.invnbr as invoice," & vbCrLf & _
"  q1.type as category" & vbCrLf & _
"from internal_sales_ht i" & vbCrLf & _
"  inner join (" & vbCrLf & _
"" & vbCrLf & _
"select distinct" & vbCrLf & _
"  i.dept," & vbCrLf & _
"  nz(dt.type_sales_analysis,'Internal Unknown') as type" & vbCrLf & _
"from internal_sales_ht i" & vbCrLf & _
"  left join dept_type dt" & vbCrLf & _
"    on i.dept = dt.dept) q1" & vbCrLf & _
"" & vbCrLf & _
"  on i.dept = q1.dept) q2"

db.Execute strsql

strsql = ""
strsql = strsql & vbCrLf & "select distinct" & vbCrLf & _
"  i.fullinvoicenbr as invoice," & vbCrLf & _
"  i.category" & vbCrLf & _
"into myop_invoice_internal" & vbCrLf & _
"from internal_sales_myop i"

db.Execute strsql

strsql = ""
strsql = strsql & vbCrLf & "update gl_sales_cogs g inner join ht_invoice_internal i" & vbCrLf & _
"  on g.invoice = i.invoice" & vbCrLf & _
"set g.category = i.category" & vbCrLf & _
"where" & vbCrLf & _
"   g.org = 'HT'"

db.Execute strsql

strsql = ""
strsql = strsql & vbCrLf & "update gl_sales_cogs g inner join myop_invoice_internal i" & vbCrLf & _
"  on g.invoice = i.invoice" & vbCrLf & _
"set g.category = i.category" & vbCrLf & _
"where" & vbCrLf & _
"   g.org = 'MYOP'"

db.Execute strsql
db.Execute "update gl_sales_cogs set category = 'External' where category = 'None'"

'Variance
'HT Sales
strsql = ""
 strsql = strsql & vbCrLf & "select" & vbCrLf & _
"  'HT' as org," & vbCrLf & _
"  q2.invoice," & vbCrLf & _
"  q2.bu," & vbCrLf & _
"  g1.category," & vbCrLf & _
"  null as type," & vbCrLf & _
"  q2.sales_gl," & vbCrLf & _
"  q2.sales_dw," & vbCrLf & _
"  q2.sales_variance" & vbCrLf & _
"from (" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  q1.invoice," & vbCrLf & _
"  q1.bu," & vbCrLf & _
"  q1.sales_gl," & vbCrLf & _
"  q1.sales_dw," & vbCrLf & _
"  q1.sales_gl - q1.sales_dw as sales_variance" & vbCrLf & _
"from (" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  g1.invoice,"
strsql = strsql & vbCrLf & "  g1.bu," & vbCrLf & _
"  g1.sales_gl," & vbCrLf & _
"  sum(iif(d1.sales_dw is null, 0, d1.sales_dw)) as sales_dw" & vbCrLf & _
"from (" & vbCrLf & _
"select" & vbCrLf & _
"  g.invoice," & vbCrLf & _
"  g.bu," & vbCrLf & _
"  sum(g.sales) as sales_gl" & vbCrLf & _
"from gl_sales_cogs g" & vbCrLf & _
"where" & vbCrLf & _
"  g.org = 'HT'" & vbCrLf & _
"group by" & vbCrLf & _
"  g.invoice," & vbCrLf & _
"  g.bu) g1" & vbCrLf & _
"" & vbCrLf & _
"  left join (" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  d.invoice," & vbCrLf & _
"  d.bu_gl,"
strsql = strsql & vbCrLf & "  sum(d.sales) as sales_dw" & vbCrLf & _
"from dw_sales_cost_invoice d" & vbCrLf & _
"where" & vbCrLf & _
"  d.system = 'NAV'" & vbCrLf & _
"group by" & vbCrLf & _
"  d.invoice," & vbCrLf & _
"  d.bu_gl) d1" & vbCrLf & _
"" & vbCrLf & _
"    on g1.invoice = d1.invoice" & vbCrLf & _
"    and g1.bu = d1.bu_gl" & vbCrLf & _
"" & vbCrLf & _
"group by" & vbCrLf & _
"  g1.invoice," & vbCrLf & _
"  g1.bu," & vbCrLf & _
"  g1.sales_gl) q1" & vbCrLf & _
"" & vbCrLf & _
"where" & vbCrLf & _
"  q1.sales_gl - q1.sales_dw <> 0) q2" & vbCrLf & _
"" & vbCrLf & _
"  left join ("
strsql = strsql & vbCrLf & "" & vbCrLf & _
"select" & vbCrLf & _
"  g.invoice," & vbCrLf & _
"  g.bu," & vbCrLf & _
"  max(g.category) as category" & vbCrLf & _
"from gl_sales_cogs g" & vbCrLf & _
"where" & vbCrLf & _
"  g.org = 'HT'" & vbCrLf & _
"group by" & vbCrLf & _
"  g.invoice," & vbCrLf & _
"  g.bu) g1" & vbCrLf & _
"" & vbCrLf & _
"    on q2.invoice = g1.invoice" & vbCrLf & _
"    and q2.bu = g1.bu"


Set qdefDAO = db.CreateQueryDef("t1", strsql)
db.Execute "insert into sales_variance_ht select * from t1"
Set qdefDAO = Nothing
DoCmd.DeleteObject acQuery, "t1"

strsql = ""
 strsql = strsql & vbCrLf & "select" & vbCrLf & _
"  'HT' as org,  " & vbCrLf & _
"  d1.invoice," & vbCrLf & _
"  d1.bu," & vbCrLf & _
"  'External' as category," & vbCrLf & _
"  null as type," & vbCrLf & _
"  0 as sales_gl," & vbCrLf & _
"  sum(dw.sales) as sales_dw," & vbCrLf & _
"  -sum(dw.sales) as sales_variance" & vbCrLf & _
"from (" & vbCrLf & _
"" & vbCrLf & _
"select distinct" & vbCrLf & _
"  d.invoice," & vbCrLf & _
"  d.bu_gl as bu" & vbCrLf & _
"from (dw_sales_cost_invoice d" & vbCrLf & _
"  left join sales_variance_ht v" & vbCrLf & _
"    on d.invoice = v.invoice" & vbCrLf & _
"    and d.bu_gl = v.bu)" & vbCrLf & _
"  left join gl_sales_cogs g" & vbCrLf & _
"    on d.invoice = g.invoice" & vbCrLf & _
"    and d.bu_gl = g.bu"
strsql = strsql & vbCrLf & "where  " & vbCrLf & _
"  d.system = 'NAV'" & vbCrLf & _
"  and v.invoice is null" & vbCrLf & _
"  and v.bu is null" & vbCrLf & _
"  and g.bu is null" & vbCrLf & _
"  and g.invoice is null" & vbCrLf & _
"  and d.sales <> 0) d1" & vbCrLf & _
"" & vbCrLf & _
"  inner join dw_sales_cost_invoice dw" & vbCrLf & _
"    on d1.invoice = dw.invoice" & vbCrLf & _
"    and d1.bu=dw.bu_gl" & vbCrLf & _
"" & vbCrLf & _
"group by" & vbCrLf & _
"  d1.invoice," & vbCrLf & _
"  d1.bu"

Set qdefDAO = db.CreateQueryDef("t1", strsql)
db.Execute "insert into sales_variance_ht select * from t1"
Set qdefDAO = Nothing
DoCmd.DeleteObject acQuery, "t1"

'HT Cost
strsql = ""
 strsql = strsql & vbCrLf & "  select " & vbCrLf & _
"  'HT' as org," & vbCrLf & _
"  q2.invoice," & vbCrLf & _
"  q2.bu," & vbCrLf & _
"  g1.category," & vbCrLf & _
"  null as type," & vbCrLf & _
"  q2.cost_gl," & vbCrLf & _
"  q2.cost_dw," & vbCrLf & _
"  q2.cost_variance" & vbCrLf & _
"from (" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  q1.invoice," & vbCrLf & _
"  q1.bu," & vbCrLf & _
"  q1.cost_gl," & vbCrLf & _
"  q1.cost_dw," & vbCrLf & _
"  q1.cost_gl - q1.cost_dw as cost_variance" & vbCrLf & _
"from (" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  g1.invoice,"
strsql = strsql & vbCrLf & "  g1.bu," & vbCrLf & _
"  g1.cost_gl," & vbCrLf & _
"  sum(iif(d1.cost_dw is null, 0, d1.cost_dw)) as cost_dw" & vbCrLf & _
"from (" & vbCrLf & _
"select" & vbCrLf & _
"  g.invoice," & vbCrLf & _
"  g.bu," & vbCrLf & _
"  sum(g.cost) as cost_gl" & vbCrLf & _
"from gl_sales_cogs g" & vbCrLf & _
"where" & vbCrLf & _
"  g.org = 'HT'" & vbCrLf & _
"group by" & vbCrLf & _
"  g.invoice," & vbCrLf & _
"  g.bu) g1" & vbCrLf & _
"" & vbCrLf & _
"  left join (" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  d.invoice," & vbCrLf & _
"  d.bu_gl,"
strsql = strsql & vbCrLf & "  sum(d.cost) as cost_dw" & vbCrLf & _
"from dw_sales_cost_invoice d" & vbCrLf & _
"where" & vbCrLf & _
"  d.system = 'NAV'" & vbCrLf & _
"group by" & vbCrLf & _
"  d.invoice," & vbCrLf & _
"  d.bu_gl) d1" & vbCrLf & _
"" & vbCrLf & _
"    on g1.invoice = d1.invoice" & vbCrLf & _
"    and g1.bu = d1.bu_gl" & vbCrLf & _
"" & vbCrLf & _
"group by" & vbCrLf & _
"  g1.invoice," & vbCrLf & _
"  g1.bu," & vbCrLf & _
"  g1.cost_gl) q1" & vbCrLf & _
"" & vbCrLf & _
"where" & vbCrLf & _
"  q1.cost_gl - q1.cost_dw <> 0) q2" & vbCrLf & _
"" & vbCrLf & _
"  left join ("
strsql = strsql & vbCrLf & "" & vbCrLf & _
"select" & vbCrLf & _
"  g.invoice," & vbCrLf & _
"  g.bu," & vbCrLf & _
"  max(g.category) as category" & vbCrLf & _
"from gl_sales_cogs g" & vbCrLf & _
"where" & vbCrLf & _
"  g.org = 'HT'" & vbCrLf & _
"group by" & vbCrLf & _
"  g.invoice," & vbCrLf & _
"  g.bu) g1" & vbCrLf & _
"" & vbCrLf & _
"    on q2.invoice = g1.invoice" & vbCrLf & _
"    and q2.bu = g1.bu"

Set qdefDAO = db.CreateQueryDef("t1", strsql)
db.Execute "insert into cost_variance_ht select * from t1"
Set qdefDAO = Nothing
DoCmd.DeleteObject acQuery, "t1"

strsql = ""
 strsql = strsql & vbCrLf & "  select " & vbCrLf & _
"  'HT' as org,  " & vbCrLf & _
"  d1.invoice," & vbCrLf & _
"  d1.bu," & vbCrLf & _
"  'External' as category," & vbCrLf & _
"  null as type," & vbCrLf & _
"  0 as cost_gl," & vbCrLf & _
"  sum(dw.cost) as cost_dw," & vbCrLf & _
"  -sum(dw.cost) as cost_variance" & vbCrLf & _
"from (" & vbCrLf & _
"" & vbCrLf & _
"select distinct" & vbCrLf & _
"  d.invoice," & vbCrLf & _
"  d.bu_gl as bu" & vbCrLf & _
"from (dw_sales_cost_invoice d" & vbCrLf & _
"  left join cost_variance_ht v" & vbCrLf & _
"    on d.invoice = v.invoice" & vbCrLf & _
"    and d.bu_gl = v.bu)" & vbCrLf & _
"  left join gl_sales_cogs g" & vbCrLf & _
"    on d.invoice = g.invoice" & vbCrLf & _
"    and d.bu_gl = g.bu"
strsql = strsql & vbCrLf & "where  " & vbCrLf & _
"  d.system = 'NAV'" & vbCrLf & _
"  and v.invoice is null" & vbCrLf & _
"  and v.bu is null" & vbCrLf & _
"  and g.bu is null" & vbCrLf & _
"  and g.invoice is null" & vbCrLf & _
"  and d.cost <> 0) d1" & vbCrLf & _
"" & vbCrLf & _
"  inner join dw_sales_cost_invoice dw" & vbCrLf & _
"    on d1.invoice = dw.invoice" & vbCrLf & _
"    and d1.bu=dw.bu_gl" & vbCrLf & _
"" & vbCrLf & _
"group by" & vbCrLf & _
"  d1.invoice," & vbCrLf & _
"  d1.bu"

Set qdefDAO = db.CreateQueryDef("t1", strsql)
db.Execute "insert into cost_variance_ht select * from t1"
Set qdefDAO = Nothing
DoCmd.DeleteObject acQuery, "t1"

'MYOP Sales
strsql = ""
 strsql = strsql & vbCrLf & "  select " & vbCrLf & _
"  'MYOP' as org," & vbCrLf & _
"  q2.invoice," & vbCrLf & _
"  q2.bu," & vbCrLf & _
"  g1.category," & vbCrLf & _
"  null as type," & vbCrLf & _
"  q2.sales_gl," & vbCrLf & _
"  q2.sales_dw," & vbCrLf & _
"  q2.sales_variance" & vbCrLf & _
"from (" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  q1.invoice," & vbCrLf & _
"  q1.bu," & vbCrLf & _
"  q1.sales_gl," & vbCrLf & _
"  q1.sales_dw," & vbCrLf & _
"  q1.sales_gl - q1.sales_dw as sales_variance" & vbCrLf & _
"from (" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  g1.invoice,"
strsql = strsql & vbCrLf & "  g1.bu," & vbCrLf & _
"  g1.sales_gl," & vbCrLf & _
"  sum(iif(d1.sales_dw is null, 0, d1.sales_dw)) as sales_dw" & vbCrLf & _
"from (" & vbCrLf & _
"select" & vbCrLf & _
"  g.invoice," & vbCrLf & _
"  g.bu," & vbCrLf & _
"  sum(g.sales) as sales_gl" & vbCrLf & _
"from gl_sales_cogs g" & vbCrLf & _
"where" & vbCrLf & _
"  g.org = 'MYOP'" & vbCrLf & _
"group by" & vbCrLf & _
"  g.invoice," & vbCrLf & _
"  g.bu) g1" & vbCrLf & _
"" & vbCrLf & _
"  left join (" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  d.invoice," & vbCrLf & _
"  d.bu_gl,"
strsql = strsql & vbCrLf & "  sum(d.sales) as sales_dw" & vbCrLf & _
"from dw_sales_cost_invoice d" & vbCrLf & _
"where" & vbCrLf & _
"  d.system = 'Vibe'" & vbCrLf & _
"group by" & vbCrLf & _
"  d.invoice," & vbCrLf & _
"  d.bu_gl) d1" & vbCrLf & _
"" & vbCrLf & _
"    on g1.invoice = d1.invoice" & vbCrLf & _
"    and g1.bu = d1.bu_gl" & vbCrLf & _
"" & vbCrLf & _
"group by" & vbCrLf & _
"  g1.invoice," & vbCrLf & _
"  g1.bu," & vbCrLf & _
"  g1.sales_gl) q1" & vbCrLf & _
"" & vbCrLf & _
"where" & vbCrLf & _
"  q1.sales_gl - q1.sales_dw <> 0) q2" & vbCrLf & _
"" & vbCrLf & _
"  left join ("
strsql = strsql & vbCrLf & "" & vbCrLf & _
"select" & vbCrLf & _
"  g.invoice," & vbCrLf & _
"  g.bu," & vbCrLf & _
"  max(g.category) as category" & vbCrLf & _
"from gl_sales_cogs g" & vbCrLf & _
"where" & vbCrLf & _
"  g.org = 'MYOP'" & vbCrLf & _
"group by" & vbCrLf & _
"  g.invoice," & vbCrLf & _
"  g.bu) g1" & vbCrLf & _
"" & vbCrLf & _
"    on q2.invoice = g1.invoice" & vbCrLf & _
"    and q2.bu = g1.bu"

Set qdefDAO = db.CreateQueryDef("t1", strsql)
db.Execute "insert into sales_variance_myop select * from t1"
Set qdefDAO = Nothing
DoCmd.DeleteObject acQuery, "t1"

strsql = ""
 strsql = strsql & vbCrLf & "select" & vbCrLf & _
"  'MYOP' as org,  " & vbCrLf & _
"  d1.invoice," & vbCrLf & _
"  d1.bu," & vbCrLf & _
"  'External' as category," & vbCrLf & _
"  null as type," & vbCrLf & _
"  0 as sales_gl," & vbCrLf & _
"  sum(dw.sales) as sales_dw," & vbCrLf & _
"  -sum(dw.sales) as sales_variance" & vbCrLf & _
"from (" & vbCrLf & _
"" & vbCrLf & _
"select distinct" & vbCrLf & _
"  d.invoice," & vbCrLf & _
"  d.bu_gl as bu" & vbCrLf & _
"from (dw_sales_cost_invoice d" & vbCrLf & _
"  left join sales_variance_ht v" & vbCrLf & _
"    on d.invoice = v.invoice" & vbCrLf & _
"    and d.bu_gl = v.bu)" & vbCrLf & _
"  left join gl_sales_cogs g" & vbCrLf & _
"    on d.invoice = g.invoice" & vbCrLf & _
"    and d.bu_gl = g.bu"
strsql = strsql & vbCrLf & "where  " & vbCrLf & _
"  d.system = 'Vibe'" & vbCrLf & _
"  and v.invoice is null" & vbCrLf & _
"  and v.bu is null" & vbCrLf & _
"  and g.bu is null" & vbCrLf & _
"  and g.invoice is null" & vbCrLf & _
"  and d.sales <> 0) d1" & vbCrLf & _
"" & vbCrLf & _
"  inner join dw_sales_cost_invoice dw" & vbCrLf & _
"    on d1.invoice = dw.invoice" & vbCrLf & _
"    and d1.bu=dw.bu_gl" & vbCrLf & _
"" & vbCrLf & _
"group by" & vbCrLf & _
"  d1.invoice," & vbCrLf & _
"  d1.bu"

Set qdefDAO = db.CreateQueryDef("t1", strsql)
db.Execute "insert into sales_variance_myop select * from t1"
Set qdefDAO = Nothing
DoCmd.DeleteObject acQuery, "t1"

'MYOP Cost
strsql = ""
 strsql = strsql & vbCrLf & "  select " & vbCrLf & _
"  'MYOP' as org," & vbCrLf & _
"  q2.invoice," & vbCrLf & _
"  q2.bu," & vbCrLf & _
"  g1.category," & vbCrLf & _
"  null as type," & vbCrLf & _
"  q2.cost_gl," & vbCrLf & _
"  q2.cost_dw," & vbCrLf & _
"  q2.cost_variance" & vbCrLf & _
"from (" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  q1.invoice," & vbCrLf & _
"  q1.bu," & vbCrLf & _
"  q1.cost_gl," & vbCrLf & _
"  q1.cost_dw," & vbCrLf & _
"  q1.cost_gl - q1.cost_dw as cost_variance" & vbCrLf & _
"from (" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  g1.invoice,"
strsql = strsql & vbCrLf & "  g1.bu," & vbCrLf & _
"  g1.cost_gl," & vbCrLf & _
"  sum(iif(d1.cost_dw is null, 0, d1.cost_dw)) as cost_dw" & vbCrLf & _
"from (" & vbCrLf & _
"select" & vbCrLf & _
"  g.invoice," & vbCrLf & _
"  g.bu," & vbCrLf & _
"  sum(g.cost) as cost_gl" & vbCrLf & _
"from gl_sales_cogs g" & vbCrLf & _
"where" & vbCrLf & _
"  g.org = 'MYOP'" & vbCrLf & _
"group by" & vbCrLf & _
"  g.invoice," & vbCrLf & _
"  g.bu) g1" & vbCrLf & _
"" & vbCrLf & _
"  left join (" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  d.invoice," & vbCrLf & _
"  d.bu_gl,"
strsql = strsql & vbCrLf & "  sum(d.cost) as cost_dw" & vbCrLf & _
"from dw_sales_cost_invoice d" & vbCrLf & _
"where" & vbCrLf & _
"  d.system = 'Vibe'" & vbCrLf & _
"group by" & vbCrLf & _
"  d.invoice," & vbCrLf & _
"  d.bu_gl) d1" & vbCrLf & _
"" & vbCrLf & _
"    on g1.invoice = d1.invoice" & vbCrLf & _
"    and g1.bu = d1.bu_gl" & vbCrLf & _
"" & vbCrLf & _
"group by" & vbCrLf & _
"  g1.invoice," & vbCrLf & _
"  g1.bu," & vbCrLf & _
"  g1.cost_gl) q1" & vbCrLf & _
"" & vbCrLf & _
"where" & vbCrLf & _
"  q1.cost_gl - q1.cost_dw <> 0) q2" & vbCrLf & _
"" & vbCrLf & _
"  left join ("
strsql = strsql & vbCrLf & "" & vbCrLf & _
"select" & vbCrLf & _
"  g.invoice," & vbCrLf & _
"  g.bu," & vbCrLf & _
"  max(g.category) as category" & vbCrLf & _
"from gl_sales_cogs g" & vbCrLf & _
"where" & vbCrLf & _
"  g.org = 'MYOP'" & vbCrLf & _
"group by" & vbCrLf & _
"  g.invoice," & vbCrLf & _
"  g.bu) g1" & vbCrLf & _
"" & vbCrLf & _
"    on q2.invoice = g1.invoice" & vbCrLf & _
"    and q2.bu = g1.bu"

Set qdefDAO = db.CreateQueryDef("t1", strsql)
db.Execute "insert into cost_variance_myop select * from t1"
Set qdefDAO = Nothing
DoCmd.DeleteObject acQuery, "t1"

strsql = ""
 strsql = strsql & vbCrLf & "  select " & vbCrLf & _
"  'MYOP' as org,  " & vbCrLf & _
"  d1.invoice," & vbCrLf & _
"  d1.bu," & vbCrLf & _
"  'External' as category," & vbCrLf & _
"  null as type," & vbCrLf & _
"  0 as cost_gl," & vbCrLf & _
"  sum(dw.cost) as cost_dw," & vbCrLf & _
"  -sum(dw.cost) as cost_variance" & vbCrLf & _
"from (" & vbCrLf & _
"" & vbCrLf & _
"select distinct" & vbCrLf & _
"  d.invoice," & vbCrLf & _
"  d.bu_gl as bu" & vbCrLf & _
"from (dw_sales_cost_invoice d" & vbCrLf & _
"  left join cost_variance_ht v" & vbCrLf & _
"    on d.invoice = v.invoice" & vbCrLf & _
"    and d.bu_gl = v.bu)" & vbCrLf & _
"  left join gl_sales_cogs g" & vbCrLf & _
"    on d.invoice = g.invoice" & vbCrLf & _
"    and d.bu_gl = g.bu"
strsql = strsql & vbCrLf & "where  " & vbCrLf & _
"  d.system = 'Vibe'" & vbCrLf & _
"  and v.invoice is null" & vbCrLf & _
"  and v.bu is null" & vbCrLf & _
"  and g.bu is null" & vbCrLf & _
"  and g.invoice is null" & vbCrLf & _
"  and d.cost <> 0) d1" & vbCrLf & _
"" & vbCrLf & _
"  inner join dw_sales_cost_invoice dw" & vbCrLf & _
"    on d1.invoice = dw.invoice" & vbCrLf & _
"    and d1.bu=dw.bu_gl" & vbCrLf & _
"" & vbCrLf & _
"group by" & vbCrLf & _
"  d1.invoice," & vbCrLf & _
"  d1.bu"

Set qdefDAO = db.CreateQueryDef("t1", strsql)
db.Execute "insert into cost_variance_myop select * from t1"
Set qdefDAO = Nothing
DoCmd.DeleteObject acQuery, "t1"

'type
db.Execute "delete from type_build"

strsql = ""
strsql = strsql & vbCrLf & "insert into type_build" & vbCrLf & _
"select  " & vbCrLf & _
"  g.invoice," & vbCrLf & _
"  g.bu," & vbCrLf & _
"  g.org," & vbCrLf & _
"  max(g.doctype) as type  " & vbCrLf & _
"from gl_sales_cogs g" & vbCrLf & _
"group by" & vbCrLf & _
"  g.invoice," & vbCrLf & _
"  g.bu," & vbCrLf & _
"  g.org"

db.Execute strsql

strsql = ""
strsql = strsql & vbCrLf & "insert into type_build" & vbCrLf & _
"select" & vbCrLf & _
"  d.invoice," & vbCrLf & _
"  d.bu_gl as bu," & vbCrLf & _
"  iif(d.system = 'NAV', 'HT', 'MYOP') as org," & vbCrLf & _
"  max(d.doctypedesc) as type" & vbCrLf & _
"from dw_sales_cost_invoice d" & vbCrLf & _
"  left join type_build t" & vbCrLf & _
"    on d.invoice = t.invoice" & vbCrLf & _
"    and d.bu_gl = t.bu" & vbCrLf & _
"where" & vbCrLf & _
"  t.invoice is null" & vbCrLf & _
"  and t.bu is null" & vbCrLf & _
"group by" & vbCrLf & _
"  d.invoice," & vbCrLf & _
"  d.bu_gl," & vbCrLf & _
"  iif(d.system = 'NAV', 'HT', 'MYOP')"

db.Execute strsql

strsql = ""
strsql = strsql & vbCrLf & "update sales_variance_ht v  left join type_build t" & vbCrLf & _
"  on v.invoice = t.invoice" & vbCrLf & _
"  and v.bu = t.bu" & vbCrLf & _
"set v.type = t.type" & vbCrLf & _
"where" & vbCrLf & _
"  t.org = 'HT'"

db.Execute strsql

strsql = ""
strsql = strsql & vbCrLf & "update sales_variance_myop v  left join type_build t" & vbCrLf & _
"  on v.invoice = t.invoice" & vbCrLf & _
"  and v.bu = t.bu" & vbCrLf & _
"set v.type = t.type" & vbCrLf & _
"where" & vbCrLf & _
"  t.org = 'MYOP'"

db.Execute strsql

strsql = ""
strsql = strsql & vbCrLf & "update cost_variance_ht v  left join type_build t" & vbCrLf & _
"  on v.invoice = t.invoice" & vbCrLf & _
"  and v.bu = t.bu" & vbCrLf & _
"set v.type = t.type" & vbCrLf & _
"where" & vbCrLf & _
"  t.org = 'HT'"

db.Execute strsql

strsql = ""
strsql = strsql & vbCrLf & "update cost_variance_myop v  left join type_build t" & vbCrLf & _
"  on v.invoice = t.invoice" & vbCrLf & _
"  and v.bu = t.bu" & vbCrLf & _
"set v.type = t.type" & vbCrLf & _
"where" & vbCrLf & _
"  t.org = 'MYOP'"

db.Execute strsql

'combine variances
strsql = ""
strsql = strsql & vbCrLf & "select" & vbCrLf & _
"  'Sales' as measure," & vbCrLf & _
"  svh.org," & vbCrLf & _
"  svh.invoice," & vbCrLf & _
"  svh.bu," & vbCrLf & _
"  svh.category," & vbCrLf & _
"  svh.type," & vbCrLf & _
"  svh.sales_gl as gl," & vbCrLf & _
"  svh.sales_dw as dw," & vbCrLf & _
"  svh.sales_variance as variance" & vbCrLf & _
"from sales_variance_ht svh" & vbCrLf & _
"" & vbCrLf & _
"union all" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  'Cost' as measure," & vbCrLf & _
"  cvh.org," & vbCrLf & _
"  cvh.invoice," & vbCrLf & _
"  cvh.bu," & vbCrLf & _
"  cvh.category,"
strsql = strsql & vbCrLf & "  cvh.type," & vbCrLf & _
"  cvh.cost_gl as gl," & vbCrLf & _
"  cvh.cost_dw as dw," & vbCrLf & _
"  cvh.cost_variance as variance" & vbCrLf & _
"from cost_variance_ht cvh" & vbCrLf & _
"" & vbCrLf & _
"union all" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  'Sales' as measure," & vbCrLf & _
"  svm.org," & vbCrLf & _
"  svm.invoice," & vbCrLf & _
"  svm.bu," & vbCrLf & _
"  svm.category," & vbCrLf & _
"  svm.type," & vbCrLf & _
"  svm.sales_gl as gl," & vbCrLf & _
"  svm.sales_dw as dw," & vbCrLf & _
"  svm.sales_variance as variance" & vbCrLf & _
"from sales_variance_myop svm" & vbCrLf & _
""
strsql = strsql & vbCrLf & "union all" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"  'Cost' as measure," & vbCrLf & _
"  cvm.org," & vbCrLf & _
"  cvm.invoice," & vbCrLf & _
"  cvm.bu," & vbCrLf & _
"  cvm.category," & vbCrLf & _
"  cvm.type," & vbCrLf & _
"  cvm.cost_gl as gl," & vbCrLf & _
"  cvm.cost_dw as dw," & vbCrLf & _
"  cvm.cost_variance as variance" & vbCrLf & _
"from cost_variance_myop cvm"

Set qdefDAO = db.CreateQueryDef("t1", strsql)
db.Execute "insert into variance_all select * from t1"
Set qdefDAO = Nothing
DoCmd.DeleteObject acQuery, "t1"

'update montefiore types
db.Execute "delete from montefiore_copier_invoice"

strsql = ""
strsql = strsql & vbCrLf & "insert into montefiore_copier_invoice" & vbCrLf & _
"select" & vbCrLf & _
"  d.custname," & vbCrLf & _
"  d.invoice," & vbCrLf & _
"  sum(d.cost) as cost" & vbCrLf & _
"from dw_sales_cost_invoice d" & vbCrLf & _
"where" & vbCrLf & _
"  d.custname in ('MPS-Montefiore Copier Program','MPS-Montefiore Mt Vernon Copiers','MPS-Montefiore New Rochelle Copiers')" & vbCrLf & _
"group by" & vbCrLf & _
"  d.custname," & vbCrLf & _
"  d.invoice"

db.Execute strsql

strsql = ""
strsql = strsql & vbCrLf & "update variance_all v inner join montefiore_copier_invoice m" & vbCrLf & _
"  on v.invoice = m.invoice" & vbCrLf & _
"set v.type = m.custname"

db.Execute strsql

'update ppi, ppc
db.Execute "update variance_all set type = left(invoice, 3) where left(invoice, 3) in ('PPI', 'PPC')"

'update known je doc numbers
strsql = ""
strsql = strsql & vbCrLf & "update variance_all v inner join je_doc_nbr j" & vbCrLf & _
"  on v.org = j.gl_company" & vbCrLf & _
"  and v.invoice = j.doc_nbr" & vbCrLf & _
"set" & vbCrLf & _
"  v.category = j.type," & vbCrLf & _
"  v.type = j.type "

db.Execute strsql

'BU Reclass Variance
On Error Resume Next
DoCmd.DeleteObject acTable, "bu_reclass"
On Error GoTo 0

strsql = ""
strsql = strsql & vbCrLf & "select" & vbCrLf & _
"   q2.measure," & vbCrLf & _
"   q2.org," & vbCrLf & _
"   q2.invoice," & vbCrLf & _
"   v1.type," & vbCrLf & _
"   'BU Reclass' as category," & vbCrLf & _
"   q2.GL," & vbCrLf & _
"   q2.DW," & vbCrLf & _
"   q2.Variance" & vbCrLf & _
"into bu_reclass" & vbCrLf & _
"from (" & vbCrLf & _
"select" & vbCrLf & _
"   'Sales' as measure," & vbCrLf & _
"   q1.org," & vbCrLf & _
"   q1.invoice," & vbCrLf & _
"   sum(q1.GL) as GL," & vbCrLf & _
"   sum(q1.DW) as DW," & vbCrLf & _
"   sum(q1.GL) -  sum(q1.DW) as Variance" & vbCrLf & _
"from (" & vbCrLf & _
"select" & vbCrLf & _
"   v.org,"
strsql = strsql & vbCrLf & "   v.invoice," & vbCrLf & _
"   sum(iif(v.measure = 'Sales', v.gl, 0)) as GL," & vbCrLf & _
"   sum(iif(v.measure = 'Sales', v.dw, 0)) as DW" & vbCrLf & _
"from variance_all v" & vbCrLf & _
"group by" & vbCrLf & _
"   v.org," & vbCrLf & _
"   v.invoice) q1" & vbCrLf & _
"group by" & vbCrLf & _
"   q1.org," & vbCrLf & _
"   q1.invoice" & vbCrLf & _
"having" & vbCrLf & _
"   (sum(q1.GL) -  sum(q1.DW)) = 0" & vbCrLf & _
"   and (abs(sum(q1.GL)) > 0 or abs(sum(q1.DW)) > 0)" & vbCrLf & _
"" & vbCrLf & _
"union all" & vbCrLf & _
"" & vbCrLf & _
"select" & vbCrLf & _
"   'Cost' as measure," & vbCrLf & _
"   q1.org," & vbCrLf & _
"   q1.invoice,"
strsql = strsql & vbCrLf & "   sum(q1.GL) as GL," & vbCrLf & _
"   sum(q1.DW) as DW," & vbCrLf & _
"   sum(q1.GL) -  sum(q1.DW) as Variance" & vbCrLf & _
"from (" & vbCrLf & _
"select" & vbCrLf & _
"   v.org," & vbCrLf & _
"   v.invoice," & vbCrLf & _
"   sum(iif(v.measure = 'Cost', v.gl, 0)) as GL," & vbCrLf & _
"   sum(iif(v.measure = 'Cost', v.dw, 0)) as DW" & vbCrLf & _
"from variance_all v" & vbCrLf & _
"group by" & vbCrLf & _
"   v.org," & vbCrLf & _
"   v.invoice) q1" & vbCrLf & _
"group by" & vbCrLf & _
"   q1.org," & vbCrLf & _
"   q1.invoice" & vbCrLf & _
"having" & vbCrLf & _
"   (sum(q1.GL) -  sum(q1.DW)) = 0" & vbCrLf & _
"   and (abs(sum(q1.GL)) > 0 or abs(sum(q1.DW)) > 0)) q2" & vbCrLf & _
"   inner join ("
strsql = strsql & vbCrLf & "      select" & vbCrLf & _
"         v.measure, " & vbCrLf & _
"         v.org," & vbCrLf & _
"         v.invoice," & vbCrLf & _
"         max(v.type) as type," & vbCrLf & _
"         max(v.category) as category " & vbCrLf & _
"      from variance_all v" & vbCrLf & _
"      group by" & vbCrLf & _
"         v.measure, " & vbCrLf & _
"         v.org," & vbCrLf & _
"         v.invoice) v1" & vbCrLf & _
"      on q2.measure = v1.measure" & vbCrLf & _
"      and q2.org = v1.org" & vbCrLf & _
"      and q2.invoice = v1.invoice"

db.Execute strsql

strsql = ""
strsql = strsql & vbCrLf & "update variance_all v inner join bu_reclass br" & vbCrLf & _
"   on v.measure = br.measure" & vbCrLf & _
"   and v.org = br.org" & vbCrLf & _
"   and v.invoice = br.invoice" & vbCrLf & _
"set v.category = br.category, v.type = br.category"

db.Execute strsql

'MYOP Seller Flex
'strsql = ""
'strsql = strsql & vbCrLf & "declare @datestart date = '" & CStr(dteStart) & "'" & vbCrLf & _
'"declare @dateend date = '" & CStr(dteEnd) & "'" & vbCrLf & _
'"" & vbCrLf & _
'"select distinct" & vbCrLf & _
'"q2.[Document No_] as doc_number" & vbCrLf & _
'"from (" & vbCrLf & _
'"select" & vbCrLf & _
'"q1.[Customer No_]," & vbCrLf & _
'"q1.[Document No_]," & vbCrLf & _
'"gl.[Posting Date]," & vbCrLf & _
'"gl.Description," & vbCrLf & _
'"sum(case when glr.level_1 = 'Sales' then -gl.amount else 0 end) as Sales," & vbCrLf & _
'"sum(case when glr.level_1 = 'Cost' then gl.amount else 0 end) as Cost" & vbCrLf & _
'"from (" & vbCrLf & _
'"select" & vbCrLf & _
'"cle.[Customer No_]," & vbCrLf & _
'"cle.[Document No_]" & vbCrLf & _
'"from NAVRep.dbo.[MYOP$Cust_ Ledger Entry] cle with(nolock)" & vbCrLf & _
'"where" & vbCrLf & _
'"cle.[Posting Date] between @datestart and @dateend"
'strsql = strsql & vbCrLf & "and cle.[Customer No_] = '90884'" & vbCrLf & _
'"and cle.[Document Type] = 2" & vbCrLf & _
'") q1" & vbCrLf & _
'"inner join NAVRep.dbo.[MYOP$G_L Entry] gl with(nolock)" & vbCrLf & _
'"on q1.[Document No_] = gl.[Document No_]" & vbCrLf & _
'"inner join tndcsql02.playground.[myop\jason.walker].gl_account_reporting glr" & vbCrLf & _
'"on gl.[G_L Account No_] = glr.gl_account" & vbCrLf & _
'"and glr.company = 'MYOP'" & vbCrLf & _
'"where" & vbCrLf & _
'"gl.[Posting Date] between @datestart and @dateend" & vbCrLf & _
'"and glr.level_1 in ('Sales', 'Cost')" & vbCrLf & _
'"group by" & vbCrLf & _
'"q1.[Customer No_]," & vbCrLf & _
'"q1.[Document No_]," & vbCrLf & _
'"gl.[Posting Date]," & vbCrLf & _
'"gl.Description" & vbCrLf & _
'") q2" & vbCrLf & _
'"where" & vbCrLf & _
'"abs(q2.Sales) > 0" & vbCrLf & _
'"or abs(q2.Cost) > 0"
'
'
'Set qdefDAO = db.QueryDefs("q_navrep_select")
'qdefDAO.SQL = strsql
'
'On Error Resume Next
'DoCmd.DeleteObject acTable, "myop_seller_flex_invoice"
'On Error GoTo 0
'
'db.Execute "select * into myop_seller_flex_invoice from q_navrep_select"
'Set qdefDAO = Nothing
'
'strsql = ""
'strsql = strsql & vbCrLf & "update variance_all v inner join myop_seller_flex_invoice msf" & vbCrLf & _
'"   on v.invoice = msf.doc_number   " & vbCrLf & _
'"set" & vbCrLf & _
'"   type = 'Seller Flex'," & vbCrLf & _
'"   category = 'Seller Flex'" & vbCrLf & _
'"where" & vbCrLf & _
'"   v.org = 'myop'"
'
'db.Execute strsql

Set db = Nothing
MsgBox "Success", vbInformation + vbSystemModal, "Complete"

End Sub
