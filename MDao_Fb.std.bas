Attribute VB_Name = "MDao_Fb"
Option Explicit

Sub CrtFb(Fb)
Dao.DBEngine.CreateDatabase Fb, dbLangGeneral
End Sub

Function DbCrt(Fb) As Database
Set DbCrt = Dao.DBEngine.CreateDatabase(Fb, dbLangGeneral)
End Function

Private Sub Z_BrwFb()
BrwFb SampFbzDutyDta
End Sub

Function DbzFb(Fb) As Database
Set DbzFb = Dao.DBEngine.OpenDatabase(Fb)
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
