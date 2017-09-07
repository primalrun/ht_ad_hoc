VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_gl_account_description_search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmd_update_Click()

Dim strsql As String
Dim qdefDAO As DAO.QueryDef
Dim db As DAO.Database

Set db = CurrentDb
Set qdefDAO = db.QueryDefs("qacctdescstring")
strsql = "select c.gl_company, c.acct, c.description from chart_of_accounts c where c.gl_company = '" & Me.cb_gl_company & "' and c.description like '*" & Me.tb_account_text_search & "*'"
qdefDAO.SQL = strsql
Set qdefDAO = Nothing
Set db = Nothing
Me.sf_account.SourceObject = "Query.qacctdescstring"

End Sub

