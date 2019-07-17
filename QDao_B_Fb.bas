Attribute VB_Name = "QDao_B_Fb"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Fb_Fbq."
Private Const Asm$ = "QDao"

Private Sub Z_WszFbq()
ShwWs WszFbq(SampFbzDutyDta, "Select * from KE24")
End Sub

Function WszFbq(Fb, Q, Optional Wsn) As Worksheet
'Set WszFbq = WszDrs(DrszFbq(Fb, Q), Wsn:=Wsn)
End Function

Function DrszQ(D As Database, Q) As Drs
DrszQ = DrszRs(Rs(D, Q))
End Function

Function DrszFbq(Fb, Q) As Drs
DrszFbq = DrszRs(Rs(Db(Fb), Q))
End Function

Function ArszFbq(Fb, Q) As AdoDb.Recordset
Set ArszFbq = CnzFb(Fb).Execute(Q)
End Function

Sub ArunFbq(Fb, Q)
CnzFb(Fb).Execute Q
End Sub


Function CrtFb(Fb, Optional IsDltFst As Boolean) As Database
If IsDltFst Then DltFfnIf Fb
Set CrtFb = Dao.DBEngine.CreateDatabase(Fb, dbLangGeneral)
End Function

Private Sub Z_BrwFb()
BrwFb SampFbzDutyDta
End Sub

Function DbzFb(Fb) As Database
Set DbzFb = Dao.DBEngine.OpenDatabase(Fb)
End Function

Function CntrNyzFb(Fb) As String()
CntrNyzFb = Itn(Db(Fb).Containers)
End Function

Function CntrItmNyzFb(Fb) As String()
Dim D As Database: Set D = Db(Fb)
Dim CntrNm
For Each CntrNm In Itn(D.Containers)
    PushIAy CntrItmNyzFb, AddPfxzAy(Itn(D.Containers(CntrNm).Documents), CntrNm & ".")
Next
End Function

Function Db(Fb) As Database
Set Db = Dao.DBEngine.OpenDatabase(Fb)
End Function

Sub EnsFb(Fb)
If Not HasFfn(Fb) Then CrtFb Fb
End Sub

Function OupTnyzFb(Fb) As String()
OupTnyzFb = OupTny(Db(Fb))
End Function

Sub AsgFbtStr(FbtStr$, OFb$, OT$)
If FbtStr = "" Then
    OFb = ""
    OT = ""
    Exit Sub
End If
AsgBrk FbtStr, "].[", OFb, OT
If FstChr(OFb) <> "[" Then Stop
If LasChr(OT) <> "]" Then Stop
OFb = RmvFstChr(OFb)
OT = RmvLasChr(OT)
End Sub

Sub DrpFbt(Fb, T)
CatzFb(Fb).Tables.Delete T
End Sub

Private Sub Z_HasFbt()
Ass HasFbt(SampFbzDutyDta, "SkuB")
End Sub


Private Sub Z_OupTnyzFb()
Dim Fb$
D OupTnyzFb(Fb)
End Sub

Private Sub Z_TnyzFb()
DmpAy TnyzFb(SampFbzDutyDta)
End Sub


Private Sub Z()
Z_BrwFb
MDao_Fb:
End Sub
