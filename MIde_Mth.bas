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

Function MdEns(Md As CodeModule, MthNm$, MthLines$) As CodeModule
Dim OldMthLines$: OldMthLines = MthLineszMd(Md, MthNm)
If OldMthLines = MthLines Then
    Debug.Print FmtQQ("MdEns: Mth(?) in Md(?) is same", MthNm, MdNm(Md))
End If
RmvMdMth Md, MthNm
Set MdEns = MdApdLines(Md, MthLines)
Debug.Print FmtQQ("MdEns: Mth(?) in Md(?) is replaced <=========", MthNm, MdNm(Md))
End Function

Private Sub Z_MthFTixAyzMth()
Dim A() As FTIx: A = MthFTIxAyzMth(Md("Md_"), "XX")
Dim J%
For J = 0 To UB(A)
    FTIxDmp A(J)
Next
End Sub





