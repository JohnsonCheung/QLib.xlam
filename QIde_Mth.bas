Attribute VB_Name = "QIde_Mth"
Option Explicit
Private Const CMod$ = "MIde_Mth."
Private Const Asm$ = "QIde"

Property Get MthKeyFny() As String()
MthKeyFny = SyzSsLin("PjNm MdNm Priority Nm Ty Mdy")
End Property

Function MthDNmSq(MthDNy$()) As Variant()
Dim O()
ReDim O(1 To Si(MthDNy) + 1, 1 To 6)
SetSqzDrv O, 1, MthKeyFny
Dim MthDNm, J&
For Each MthDNm In MthDNy
    SetSqzDrv O, J + 2, Split(MthDNm, ":")
    J = J + 1
Next
MthDNmSq = O
End Function

Function MdLinesAyzMth(A As CodeModule, MthNm$) As MdLines()
Dim Ix&, S$(): S = Src(A)
Dim StartLine&, Count&, I
For Each I In Itr(MthIxAyzNm(S, MthNm))
    Ix = I
    StartLine = MthTopRmkIx(S, Ix)
    Count = MthToIx(S, Ix) - StartLine + 1
    PushObj MdLinesAyzMth, MdLines(StartLine, A.Lines(StartLine, Count), Ix)
Next
End Function

Sub RplMthByDicInMd(Md As CodeModule, MthNm$, ByLines$)
Dim Ix&: Ix = MthIx(Src(Md), MthNm)
RmvMth Md, MthNm
If Ix = -1 Then
    Md.AddFromString ByLines
Else
    Md.InsertLines Ix + 1, ByLines
End If
End Sub

Private Sub Z()
Z_MthFTixAyzMth
MIde__Mth:
End Sub

Sub EnsMdLines(Md As CodeModule, MthNm$, MthLines$)
Dim OldMthLines$: OldMthLines = MthLinesByMdMth(Md, MthNm)
If OldMthLines = MthLines Then
    Debug.Print FmtQQ("EnsMd: Mth(?) in Md(?) is same", MthNm, MdNm(Md))
End If
RmvMdMth Md, MthNm
ApdLines Md, MthLines
Debug.Print FmtQQ("EnsMd: Mth(?) in Md(?) is replaced <=========", MthNm, MdNm(Md))
End Sub

Private Sub Z_MthFTixAyzMth()
Dim A() As FTIx: A = MthFTIxAyzMth(Md("Md_"), "XX")
Dim J%
For J = 0 To UB(A)
    FTIxDmp A(J)
Next
End Sub





