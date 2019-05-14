Attribute VB_Name = "QIde_Mth_Lines"
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Mth_Lines."

'aaa
Private Property Get XX1()

End Property

'BB
Private Property Let XX1(V)

End Property
Function PMthnyzP(P As VBProject) As String()

End Function
Function MthLineszPum(PMthn)

End Function

Function MthLineszPP$(P As VBProject, PMthn)
Dim B$(): B = ModNyzPum(PMthn)
If Si(B) <> 1 Then
    Thw CSub, "Should be 1 module found", "PMthn [#Mod having PMthn] ModNy-Found", PMthn, Si(B), B
End If
MthLineszPP = MthLineszSP(SrczMdn(B(0)), PMthn, WiTopRmk:=True)
End Function
'
Function MthLineszSP$(Src$(), PMthn, Optional WiTopRmk As Boolean)

End Function
'
Property Get CMthLines$() 'Cur
CMthLines = MthLineszMN(CMd, CurMthn, WiTopRmk:=True)
End Property

Sub VcMthLinesAyP()
Vc FmtLinesAy(MthLinesAyP(WiTopRmk:=True))
End Sub
Function MthLinesAyP(Optional WiTopRmk As Boolean) As String()
MthLinesAyP = MthLinesAyzP(CPj, WiTopRmk)
End Function

Function MthLinesAyzP(P As VBProject, Optional WiTopRmk As Boolean) As String()
Dim I
For Each I In MdItr(P)
    PushIAy MthLinesAyP, MthLinesAyzM(CvMd(I), WiTopRmk)
Next
End Function

Function MthLinesAyzM(A As CodeModule, Optional WiTopRmk As Boolean) As String()
MthLinesAyzM = MthLinesAyzS(Src(A), WiTopRmk)
End Function

Function MthLineszSIW$(Src$(), MthIx, Optional WiTopRmk As Boolean)
MthLineszSIW = JnCrLf(AywFEIx(Src, MthFEIxzSIW(Src, MthIx, WiTopRmk)))
End Function

Function MthLinesAyzS(Src$(), Optional WiTopRmk As Boolean) As String()
Dim Ix
For Each Ix In Itr(MthIxy(Src))
    PushI MthLinesAyzS, MthLineszSIW(Src, Ix, WiTopRmk)
Next
End Function
Function EOneMdzPM(P As VBProject, Mthn) As CodeModule
Stop '
End Function

Function MthLineszPN$(P As VBProject, Mthn, Optional WiTopRmk As Boolean)
MthLineszPN = MthLineszMN(EOneMdzPM(P, Mthn), Mthn)
End Function

Function MthLineszN$(Mthn, Optional WiTopRmk As Boolean)
MthLineszN = MthLineszPN(CPj, Mthn, WiTopRmk)
End Function


Function MthLineszMN$(Md As CodeModule, Mthn, Optional WiTopRmk As Boolean)
MthLineszMN = MthLineszSN(Src(Md), Mthn, WiTopRmk)
End Function

Function MthLineszMTN$(Md As CodeModule, ShtMthTy$, Mthn, Optional WiTopRmk As Boolean)
Dim S$(): S = Src(Md)
Dim Ix&: Ix = MthIxzSTN(S, ShtMthTy, Mthn)
MthLineszMTN = MthLineszSIW(S, Ix, WiTopRmk)
End Function

Function MthLineszSI$(Src$(), MthIx, Optional WiTopRmk As Boolean)
Dim TopLy$(): TopLy = TopRmkLyzSIW(Src, MthIx, WiTopRmk)
Dim EIx&:       EIx = MthEIx(Src, MthIx)
Dim MthLy$(): MthLy = AywFT(Src, MthIx, EIx)
MthLineszSI = JnCrLf(Sy(TopLy, MthLy))
End Function

Function MthLinzSTN$(Src$(), ShtMthTy$, Mthn)
MthLinzSTN = Src(MthIxzSTN(Src, ShtMthTy, Mthn))
End Function

Function MthLineszSN$(Src$(), Mthn, Optional WiTopRmk As Boolean)
Dim Ix, O$()
For Each Ix In Itr(MthIxyzSN(Src, Mthn))
    PushI O, MthLineszSIW(Src, Ix, WiTopRmk)
Next
MthLineszSN = JnDblCrLf(O)
End Function

Function MthLineszSTN$(Src$(), ShtMthTy$, Mthn, Optional WiTopRmk As Boolean)
Dim Ix&: Ix = MthIxzSTN(Src, ShtMthTy, Mthn)
MthLineszSTN = MthLineszSIW(Src, Ix, WiTopRmk)
End Function

