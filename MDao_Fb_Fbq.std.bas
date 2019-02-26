Attribute VB_Name = "MDao_Fb_Fbq"
Option Explicit

Private Sub Z_WszFbq()
WszFbq SampFbzDutyDta, "Select * from KE24", Vis:=True
End Sub

Function WszFbq(Fb, Sql, Optional Wsn$, Optional Vis As Boolean) As Worksheet
Set WszFbq = WszDrs(DrszFbq(Fb, Sql), Wsn:=Wsn, Vis:=Vis)
End Function

Function DrszDbq(A As Database, Q) As DRs
Set DrszDbq = DrszRs(Rs(A, Q))
End Function

Function DrszFbq(Fb, Q) As DRs
Set DrszFbq = DrszRs(Rs(Db(Fb), Q))
End Function

Function ArszFbq(A, Sql) As ADODB.Recordset
Set ArszFbq = CnzFb(A).Execute(Sql)
End Function

Sub RunFbq(A, Sql)
CnzFb(A).Execute Sql
End Sub

