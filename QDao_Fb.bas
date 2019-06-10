Attribute VB_Name = "QDao_Fb"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Fb."
Private Const Asm$ = "QDao"

Function CrtFb(FB, Optional IsDltFst As Boolean) As Database
If IsDltFst Then DltFfnIf FB
Set CrtFb = Dao.DBEngine.CreateDatabase(FB, dbLangGeneral)
End Function

Private Sub Z_BrwFb()
BrwFb SampFbzDutyDta
End Sub

Function DbzFb(FB) As Database
Set DbzFb = Dao.DBEngine.OpenDatabase(FB)
End Function

Function CntrNyzFb(FB) As String()
CntrNyzFb = Itn(Db(FB).Containers)
End Function

Function CntrItmNyzFb(FB) As String()
Dim D As Database: Set D = Db(FB)
Dim CntrNm
For Each CntrNm In Itn(D.Containers)
    PushIAy CntrItmNyzFb, AddPfxzAy(Itn(D.Containers(CntrNm).Documents), CntrNm & ".")
Next
End Function

Function Db(FB) As Database
Set Db = Dao.DBEngine.OpenDatabase(FB)
End Function

Sub EnsFb(FB)
If Not HasFfn(FB) Then CrtFb FB
End Sub

Function OupTnyzFb(FB) As String()
OupTnyzFb = OupTny(Db(FB))
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

Sub DrpFbt(FB, T)
CatzFb(FB).Tables.Delete T
End Sub

Private Sub ZZ_HasFbt()
Ass HasFbt(SampFbzDutyDta, "SkuB")
End Sub


Private Sub ZZ_OupTnyzFb()
Dim FB$
D OupTnyzFb(FB)
End Sub

Private Sub ZZ_TnyzFb()
DmpAy TnyzFb(SampFbzDutyDta)
End Sub


Private Sub ZZ()
Z_BrwFb
MDao_Fb:
End Sub
