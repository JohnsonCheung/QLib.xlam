Attribute VB_Name = "QIde_Mth"
Option Explicit
Private Const CMod$ = "MIde_Mth."
Private Const Asm$ = "QIde"

Property Get MthKeyFny() As String()
MthKeyFny = SyzSS("Pjn Mdn Priority Nm Ty Mdy")
End Property

Function SqzMthDNy(MthDNy$()) As Variant()
Dim O()
ReDim O(1 To Si(MthDNy) + 1, 1 To 6)
SetSqzDrv O, 1, MthKeyFny
Dim MthDn, J&
For Each MthDn In MthDNy
    SetSqzDrv O, J + 2, Split(MthDn, ":")
    J = J + 1
Next
SqzMthDNy = O
End Function

Sub RplMthzMNL(Md As CodeModule, Mthn, ByLines$)
Dim Ix&: Ix = FstMthIx(Src(Md), Mthn)
RmvMth Md, Mthn
If Ix = -1 Then
    Md.AddFromString ByLines
Else
    Md.InsertLines Ix + 1, ByLines
End If
End Sub

Private Sub ZZ()
MIde__Mth:
End Sub

Sub EnsLines(Md As CodeModule, Mthn, MthLines$)
Dim OldMthLines$: OldMthLines = MthLineszMN(Md, Mthn)
If OldMthLines = MthLines Then
    Debug.Print FmtQQ("EnsMd: Mth(?) in Md(?) is same", Mthn, Mdn(Md))
End If
RmvMthzMN Md, Mthn
ApdLines Md, MthLines
Debug.Print FmtQQ("EnsMd: Mth(?) in Md(?) is replaced <=========", Mthn, Mdn(Md))
End Sub



