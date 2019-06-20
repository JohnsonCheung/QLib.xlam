Attribute VB_Name = "QIde_Mth_Lines"
Option Compare Text
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Mth_Lines."

'aaa
Private Property Get XX1()

End Property

'BB
Private Property Let XX1(V)

End Property
Function PMthNyzP(P As VBProject) As String()

End Function
Function MthLzPum(PMthn)

End Function

Function MthLzPP$(P As VBProject, PMthn)
Dim B$(): B = ModNyzPum(PMthn)
If Si(B) <> 1 Then
    Thw CSub, "Should be 1 module found", "PMthn [#Mod having PMthn] ModNy-Found", PMthn, Si(B), B
End If
MthLzPP = MthLzSP(SrczMdn(B(0)), PMthn)
End Function
'
Function MthLzSP$(Src$(), PMthn)

End Function
'
Property Get CMthL$() 'Cur
CMthL = MthLzM(CMd, CMthn)
End Property

Sub VcMthLAyP()
Vc FmtLinesAy(MthLAyP)
End Sub
Function MthLAyP() As String()
MthLAyP = MthLAyzP(CPj)
End Function

Function MthLAyzP(P As VBProject) As String()
Dim I
For Each I In MdItr(P)
    PushIAy MthLAyP, MthLAyzM(CvMd(I))
Next
End Function

Function MthLAyzM(M As CodeModule) As String()
MthLAyzM = MthLAyzS(Src(M))
End Function

Function MthLAyzS(Src$()) As String()
Dim Ix
For Each Ix In Itr(MthIxy(Src))
    PushI MthLAyzS, MthLzSI(Src, Ix)
Next
End Function
Function MdzMthn(P As VBProject, Mthn) As CodeModule
Dim C As VBComponent, O As CodeModule
For Each C In P.VBComponents
    If HasEle(PMthNyzM(C.CodeModule), Mthn) Then
        If Not IsNothing(O) Then Thw CSub, FmtQQ("Mthn fnd in 2 or more md: [?] & [?]", Mdn(O), C.Name)
        Set O = C.CodeModule
    End If
Next
If IsNothing(O) Then Thw CSub, "Mthn not fnd in any codemodule of given pj", "Pj Mthn", "P.Name,Mthn"
End Function

Function MthLzPN$(P As VBProject, Mthn)
MthLzPN = MthLzM(MdzMthn(P, Mthn), Mthn)
End Function

Function MthLzN$(Mthn)
MthLzN = MthLzPN(CPj, Mthn)
End Function

Function MthLzM$(M As CodeModule, Mthn)
MthLzM = MthLzSN(Src(M), Mthn)
End Function

Function MthLyzM(M As CodeModule, Mthn) As String()
MthLyzM = SplitCrLf(MthLzM(M, Mthn))
End Function

Function MthLzMTN$(Md As CodeModule, ShtMthTy$, Mthn)
Dim S$(): S = Src(Md)
Dim Ix&: Ix = MthIxzSTN(S, ShtMthTy, Mthn)
MthLzMTN = MthLzSI(S, Ix)
End Function

Function MthLzSI$(Src$(), MthIx)
Dim EIx&:       EIx = EndLix(Src, MthIx)
Dim MthLy$(): MthLy = AywFT(Src, MthIx, EIx)
MthLzSI = JnCrLf(MthLy)
End Function

Function MthLinzSTN$(Src$(), ShtMthTy$, Mthn)
MthLinzSTN = Src(MthIxzSTN(Src, ShtMthTy, Mthn))
End Function

Function MthLzSN$(Src$(), Mthn)
Dim Ix, O$()
For Each Ix In Itr(MthIxyzSN(Src, Mthn))
    PushI O, MthLzSI(Src, Ix)
Next
MthLzSN = JnDblCrLf(O)
End Function

Function MthLzSTN$(Src$(), ShtMthTy$, Mthn)
Dim Ix&: Ix = MthIxzSTN(Src, ShtMthTy, Mthn)
MthLzSTN = MthLzSI(Src, Ix)
End Function

