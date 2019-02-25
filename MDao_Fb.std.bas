Attribute VB_Name = "MDao_Fb"
Option Explicit

Sub CrtFb(Fb)
DAO.DBEngine.CreateDatabase Fb, dbLangGeneral
End Sub

Function DbCrt(Fb) As Database
Set DbCrt = DAO.DBEngine.CreateDatabase(Fb, dbLangGeneral)
End Function

Private Sub Z_BrwFb()
BrwFb SampFbzDutyDta
End Sub


Function Db(Fb) As Database
Set Db = DAO.DBEngine.OpenDatabase(Fb)
End Function

Sub EnsFb(Fb)
If Not HasFfn(Fb) Then CrtFb Fb
End Sub


Function OupTnyzFb(Fb) As String()
OupTnyzFb = OupTny
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

Sub DrpvFbt(Fb$, T$)
CatzFb(Fb).Tables.Delete T
End Sub

Private Sub ZZ_HasFbt()
Ass HasFbt(SampFbzDutyDta, "SkuB")
End Sub


Private Sub ZZ_OupTnyzFb()
Dim Fb$
D OupTnyzFb(Fb)
End Sub

Private Sub ZZ_TnyzFb()
DmpAy TnyzFb(SampFbzDutyDta)
End Sub


Private Sub Z()
Z_BrwFb
MDao_Fb:
End Sub
