Attribute VB_Name = "MIde_Mth"
Option Explicit

Property Get MthKeyDrFny() As String()
MthKeyDrFny = SySsl("PjNm MdNm Priority Nm Ty Mdy")
End Property

Function MthDNySq(A$()) As Variant()
Dim O(), J%
ReDim O(1 To Sz(A) + 1, 1 To 6)
SetSqrzDr O, 1, MthKeyDrFny
For J = 0 To UB(A)
    SetSqrzDr O, J + 2, Split(A(J), ":")
Next
MthDNySq = O
End Function

Function MthDNmzLin$(MthLin)
MthDNmzLin = MthDNmzMthNm3(MthNm3(MthLin))
End Function

Sub RplMth(Md As CodeModule, MthDNm, ByLines)
RmvMdMth Md, MthDNm
AppLines Md, ByLines
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





