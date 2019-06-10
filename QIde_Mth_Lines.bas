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
Function PMthnyzP(P As VBProject) As String()

End Function
Function MthLineszPum(PMthn)

End Function

Function MthLineszPP$(P As VBProject, PMthn)
Dim B$(): B = ModNyzPum(PMthn)
If Si(B) <> 1 Then
    Thw CSub, "Should be 1 module found", "PMthn [#Mod having PMthn] ModNy-Found", PMthn, Si(B), B
End If
MthLineszPP = MthLineszSP(SrczMdn(B(0)), PMthn)
End Function
'
Function MthLineszSP$(Src$(), PMthn)

End Function
'
Property Get CMthLines$() 'Cur
CMthLines = MthLineszM(CMd, CMthn)
End Property

Sub VcMthLinesAyP()
Vc FmtLinesAy(MthLinesAyP)
End Sub
Function MthLinesAyP() As String()
MthLinesAyP = MthLinesAyzP(CPj)
End Function

Function MthLinesAyzP(P As VBProject) As String()
Dim I
For Each I In MdItr(P)
    PushIAy MthLinesAyP, MthLinesAyzM(CvMd(I))
Next
End Function

Function MthLinesAyzM(M As CodeModule) As String()
MthLinesAyzM = MthLinesAyzS(Src(M))
End Function

Function MthLinesAyzS(Src$()) As String()
Dim Ix
For Each Ix In Itr(MthIxy(Src))
    PushI MthLinesAyzS, MthLineszSI(Src, Ix)
Next
End Function
Function MdzMthn(P As VBProject, Mthn) As CodeModule
Dim C As VBComponent, O As CodeModule
For Each C In P.VBComponents
    If HasEle(PMthnyzM(C.CodeModule), Mthn) Then
        If Not IsNothing(O) Then Thw CSub, FmtQQ("Mthn fnd in 2 or more md: [?] & [?]", Mdn(O), C.Name)
        Set O = C.CodeModule
    End If
Next
If IsNothing(O) Then Thw CSub, "Mthn not fnd in any codemodule of given pj", "Pj Mthn", "P.Name,Mthn"
End Function

Function MthLineszPN$(P As VBProject, Mthn)
MthLineszPN = MthLineszM(MdzMthn(P, Mthn), Mthn)
End Function

Function MthLineszN$(Mthn)
MthLineszN = MthLineszPN(CPj, Mthn)
End Function

Function MthLineszM$(M As CodeModule, Mthn)
MthLineszM = MthLineszSN(Src(M), Mthn)
End Function

Function MthLyzM(M As CodeModule, Mthn) As String()
MthLyzM = SplitCrLf(MthLineszM(M, Mthn))
End Function

Function MthLineszMTN$(Md As CodeModule, ShtMthTy$, Mthn)
Dim S$(): S = Src(Md)
Dim Ix&: Ix = MthIxzSTN(S, ShtMthTy, Mthn)
MthLineszMTN = MthLineszSI(S, Ix)
End Function

Function MthLineszSI$(Src$(), MthIx)
Dim EIx&:       EIx = EndLix(Src, MthIx)
Dim MthLy$(): MthLy = AywFT(Src, MthIx, EIx)
MthLineszSI = JnCrLf(MthLy)
End Function

Function MthLinzSTN$(Src$(), ShtMthTy$, Mthn)
MthLinzSTN = Src(MthIxzSTN(Src, ShtMthTy, Mthn))
End Function

Function MthLineszSN$(Src$(), Mthn)
Dim Ix, O$()
For Each Ix In Itr(MthIxyzSN(Src, Mthn))
    PushI O, MthLineszSI(Src, Ix)
Next
MthLineszSN = JnDblCrLf(O)
End Function

Function MthLineszSTN$(Src$(), ShtMthTy$, Mthn)
Dim Ix&: Ix = MthIxzSTN(Src, ShtMthTy, Mthn)
MthLineszSTN = MthLineszSI(Src, Ix)
End Function

