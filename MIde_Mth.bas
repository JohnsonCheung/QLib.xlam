Attribute VB_Name = "MIde_Mth"
Option Explicit

Property Get MthKeyDrFny() As String()
MthKeyDrFny = SySsl("PjNm MdNm Priority Nm Ty Mdy")
End Property

Function MthQDNySq(A$()) As Variant()
Dim O(), J%
ReDim O(1 To Sz(A) + 1, 1 To 6)
SetSqrzDr O, 1, MthKeyDrFny
For J = 0 To UB(A)
    SetSqrzDr O, J + 2, Split(A(J), ":")
Next
MthQDNySq = O
End Function

Function MthDNm$(MthLin)
MthDNm = MthDNmMthNm3(MthNm3(MthLin))
End Function

Sub RplMthLines(Md As CodeModule, MthNm$, ByLines$)
RmvMdMth Md, MthNm
AppMdLines Md, ByLines
End Sub

Private Sub Z()
Z_MthFTIxAyMdMth
MIde__Mth:
End Sub

Function MdEnsMth(Md As CodeModule, MthNm$, MthLines$)
Dim OldMthLines$: OldMthLines = MthLineszMd(Md, MthNm)
If OldMthLines = MthLines Then
    Debug.Print FmtQQ("MdEnsMth: Mth(?) in Md(?) is same", MthNm, MdNm(Md))
End If
RmvMdMth Md, MthNm
AppMdLines Md, MthLines
Debug.Print FmtQQ("MdEnsMth: Mth(?) in Md(?) is replaced <=========", MthNm, MdNm(Md))
End Function


Private Sub Z_MthFTIxAyMdMth()
Dim A() As FTIx: A = MthFTIxAyMdMth(Md("Md_"), "XX")
Dim J%
For J = 0 To UB(A)
    FTIxDmp A(J)
Next
End Sub
Function MthQDNy() As String()
MthQDNy = MthQDNyMd(CurMd)
End Function
Function MthQDNyMd(A As CodeModule) As String()
'MthQDNyMd = AyAddPfx(MthDDNySrc(Src(A)), MdQNmMd(A) & ".")
End Function





