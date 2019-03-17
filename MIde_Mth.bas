Attribute VB_Name = "MIde_Mth"
Option Explicit

Property Get MthKeyDrFny() As String()
MthKeyDrFny = SySsl("PjNm MdNm Priority Nm Ty Mdy")
End Property

Function MthDNySq(A$()) As Variant()
Dim O(), J%
ReDim O(1 To Si(A) + 1, 1 To 6)
SetSqrzDr O, 1, MthKeyDrFny
For J = 0 To UB(A)
    SetSqrzDr O, J + 2, Split(A(J), ":")
Next
MthDNySq = O
End Function

Function MdLinesAyzMth(A As CodeModule, MthNm) As MdLines()
Dim Ix, S$(): S = Src(A)
Dim StartLine&, Count&
For Each Ix In Itr(MthIxAyzNm(S, MthNm))
    StartLine = MthTopRmkIx(S, Ix)
    Count = MthToIx(S, Ix) - StartLine + 1
    PushObj MdLinesAyzMth, MdLines(StartLine, A.Lines(StartLine, Count), Ix)
Next
End Function

Sub RplMth(Md As CodeModule, MthNm, ByLines)
Dim M() As MdLines: M = MdLinesAyzMth(Md, MthNm)
Select Case Si(M)
Case 0: AppLines Md, ByLines: InfLin CSub, "MthNm is added", "Md Mth MthLinCntSz", MdNm(Md), MthNm, CntSzStrzLines(ByLines)
Case 1: RplMdLines Md, M(0), ByLines, "MthLines"
Case 2: RplMdLines Md, M(0), ByLines, "MthLines": Md.DeleteLines M(1).StartLine, M(1).Count
Case Else: Thw CSub, "Er in MdLinesAyzMth.  It should return Si of 0,1 or 2", "But-Now-It-Return-Si", Si(M)
End Select
End Sub

Private Sub Z()
Z_MthFTixAyzMth
MIde__Mth:
End Sub

Function EnsMth(Md As CodeModule, MthNm$, MthLines$)
Dim OldMthLines$: OldMthLines = MthLineszMd(Md, MthNm)
If OldMthLines = MthLines Then
    Debug.Print FmtQQ("EnsMth: Mth(?) in Md(?) is same", MthNm, MdNm(Md))
End If
RmvMdMth Md, MthNm
AppLines Md, MthLines
Debug.Print FmtQQ("EnsMth: Mth(?) in Md(?) is replaced <=========", MthNm, MdNm(Md))
End Function

Private Sub Z_MthFTixAyzMth()
Dim A() As FTIx: A = MthFTIxAyzMth(Md("Md_"), "XX")
Dim J%
For J = 0 To UB(A)
    FTIxDmp A(J)
Next
End Sub

Function MthDNyzMd(A As CodeModule) As String()
MthDNyzMd = AyAddPfx(MthDNyzSrc(Src(A)), MdQNmzMd(A) & ".")
End Function





