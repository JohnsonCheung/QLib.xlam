Attribute VB_Name = "QIde_Mth"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth."
Private Const Asm$ = "QIde"

Property Get MthKeyFny() As String()
MthKeyFny = SyzSS("Pjn Mdn Priority Nm Ty Mdy")
End Property

Function SqzMthDNy(MthDNy$()) As Variant()
Dim O()
ReDim O(1 To Si(MthDNy) + 1, 1 To 6)
SetSqr O, 1, MthKeyFny
Dim MthDn, J&
For Each MthDn In MthDNy
    SetSqr O, J + 2, Split(MthDn, ":")
    J = J + 1
Next
SqzMthDNy = O
End Function

Function RplMth(M As CodeModule, Mthn, NewL$) As Boolean
'Ret True if Rplaced
Dim OldL$, Lno&
Lno = MthLnozMM(M, Mthn)
If HasMthzM(M, Mthn) Then
    OldL = MthLineszMN(M, Mthn)
    If OldL <> NewL Then
        RplMth = True
        RmvMth M, Mthn
        M.InsertLines Lno, NewL '<==
    End If
Else
    RplMth = True
    M.AddFromString NewL '<===
End If
End Function

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



