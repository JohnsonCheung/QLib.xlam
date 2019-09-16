Attribute VB_Name = "MxMthDic"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxMthDic."
':SDiMdnqSrc$ = "SDiMdnqSrc is Srtd-Mdn-To-SrcL."

Function DiMthnqLineszP(P As VBProject) As Dictionary
Dim C As VBComponent
For Each C In P.VBComponents
    PushDic DiMthnqLineszP, DiMthnqLineszM(C.CodeModule)
Next
End Function

Private Sub Z_SDiMthnqLinesM()
B SDiMthnqLinesM
End Sub
Function SDiMthnqLinesM() As Dictionary
Set SDiMthnqLinesM = SDiMthnqLineszM(CMd)
End Function
Function SDiMthnqLineszM(M As CodeModule) As Dictionary
Set SDiMthnqLineszM = SrtDic(DiMthnqLineszM(M))
End Function

Private Sub Z_DiMthnqLineszP()
Dim A As Dictionary: Set A = DiMthnqLineszP(CPj)
Ass IsDicLines(A) '
Vc A
End Sub

Private Sub Z_DiMthnqLinesM()
B DiMthnqLinesM
End Sub

Function DiMthnqLinesP() As Dictionary
Set DiMthnqLinesP = DiMthnqLineszP(CPj)
End Function

Function DiMthnqLinesM() As Dictionary
Set DiMthnqLinesM = DiMthnqLineszM(CMd)
End Function

Function DiMthnqLines(Src$(), Optional Mdn$, Optional ExlDcl As Boolean) As Dictionary 'Key is MthDn, Val is MthLWiTopRmk
Set DiMthnqLines = New Dictionary
Dim P$: If Mdn <> "" Then P = Mdn & "."
With DiMthnqLines
    If Not ExlDcl Then .Add P & "*Dcl", Dcl(Src)
    Dim Ix: For Each Ix In MthIxItr(Src)
        Dim Dn$:       Dn = MthDnzL(Src(Ix))
        Dim Lines$: Lines = MthLzIx(Src, Ix)
        .Add P & Dn, Lines
    Next
End With
End Function

Function DiMthnqLineszM(M As CodeModule, Optional ExlDcl As Boolean) As Dictionary
Set DiMthnqLineszM = DiMthnqLines(Src(M), Mdn(M), ExlDcl)
End Function

Function LineszJnLinesItr$(LinesItr, Optional Sep$ = vbCrLf)
LineszJnLinesItr = Jn(IntozItr(EmpSy, LinesItr), Sep)
End Function

Function SDiMthnqLineszP(P As VBProject) As Dictionary
Set SDiMthnqLineszP = SrtDic(DiMthnqLineszP(P))
End Function

Function SDiMthnqLinesP() As Dictionary
Set SDiMthnqLinesP = SDiMthnqLineszP(CPj)
End Function

Function SrcLzM$(M As CodeModule)
If M.CountOfLines > 0 Then
    SrcLzM = M.Lines(1, M.CountOfLines)
End If
End Function

Sub BrwSrtRptzM(M As CodeModule)
Dim Old$: Old = SrcLzM(M)
Dim NewLines$: NewLines = SrtdSrclzM(M)
Dim O$: O = IIf(Old = NewLines, "(Same)", "<====Diff")
Debug.Print Mdn(M), O
End Sub


Private Sub SrtPj(P As VBProject)
BackupFfn Pjf(P)
Dim C As VBComponent
For Each C In P.VBComponents
    SrtMd C.CodeModule
Next
End Sub


Private Sub Z_Dcl_BefAndAft_Srt()
Const Mdn$ = "VbStrRe"
Dim A$() ' Src
Dim B$() ' Src->Srt
Dim A1$() 'Src->Dcl
Dim B1$() 'Src->Src->Dcl
A = Src(Md(Mdn))
B = SrtSrc(A)
A1 = Dcl(A)
B1 = Dcl(B)
Stop
End Sub