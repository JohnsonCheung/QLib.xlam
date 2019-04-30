Attribute VB_Name = "MDao_Fb_Fbq"
Option Explicit

Private Sub Z_WszFbq()
VisWs WszFbq(SampFbzDutyDta, "Select * from KE24")
End Sub

Function WszFbq(Fb$, Q$, Optional Wsn$) As Worksheet
Set WszFbq = WszDrs(DrszFbq(Fb$, Q), Wsn:=Wsn)
End Function

Function DrszQ(A As Database, Q$) As Drs
DrszQ = DrszRs(Rs(A, Q))
End Function

Function DrszFbq(Fb$, Q$) As Drs
DrszFbq = DrszRs(Rs(Db(Fb$), Q))
End Function

Function ArszFbq(Fb$, Q$) As AdoDb.Recordset
Set ArszFbq = CnzFb(Fb).Execute(Q)
End Function

Sub ArunFbq(Fb$, Q$)
CnzFb(Fb).Execute Q
End Sub

