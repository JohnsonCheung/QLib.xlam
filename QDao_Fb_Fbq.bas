Attribute VB_Name = "QDao_Fb_Fbq"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Fb_Fbq."
Private Const Asm$ = "QDao"

Private Sub Z_WszFbq()
ShwWs WszFbq(SampFbzDutyDta, "Select * from KE24")
End Sub

Function WszFbq(FB, Q, Optional Wsn) As Worksheet
'Set WszFbq = WszDrs(DrszFbq(Fb, Q), Wsn:=Wsn)
End Function

Function DrszQ(A As Database, Q) As Drs
DrszQ = DrszRs(Rs(A, Q))
End Function

Function DrszFbq(FB, Q) As Drs
DrszFbq = DrszRs(Rs(Db(FB), Q))
End Function

Function ArszFbq(FB, Q) As AdoDb.Recordset
Set ArszFbq = CnzFb(FB).Execute(Q)
End Function

Sub ArunFbq(FB, Q)
CnzFb(FB).Execute Q
End Sub

