VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_gl_account_search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cb_account_GotFocus()

strsql = "select c.acct from chart_of_accounts c where c.gl_company = '" & Me.cb_gl_company & "'"
Me.cb_account.RowSource = strsql

End Sub


Private Sub cmd_update_Click()

Dim strsql As String
Dim qdefDAO As DAO.QueryDef
Dim db As DAO.Database

Set db = CurrentDb
Set qdefDAO = db.QueryDefs("qacct")
strsql = "select c.gl_company, c.acct, c.description from chart_of_accounts c where c.gl_company = '" & Me.cb_gl_company & "' and c.acct = '" & Me.cb_account & "'"
qdefDAO.SQL = strsql
Set qdefDAO = Nothing
Set db = Nothing
Me.sf_account.SourceObject = "Query.qacct"

End Sub
