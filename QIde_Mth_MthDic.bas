Attribute VB_Name = "QIde_Mth_MthDic"
Option Compare Text
Option Explicit
Const CMod$ = "MIde_Mth_Dic."
Const Asm$ = "QIde"
':SDiMdnqSrc$ = "SDiMdnqSrc is Srted-Mdn-To-SrcL."

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

Function SSrcL$(Src$())
':SSrcL :SrcL #Sorted-SrcLines#
SSrcL = JnStrDic(SrtDic(DiMthnqLines(Src)), vb2CrLf)
End Function
Function SSrcLM$()
SSrcLM = SSrcLzM(CMd)
End Function

Function SSrcLzM$(M As CodeModule)
SSrcLzM = SSrcL(Src(M))
End Function

Function SrcLzM$(M As CodeModule)
If M.CountOfLines > 0 Then
    SrcLzM = M.Lines(1, M.CountOfLines)
End If
End Function

Sub BrwSrtRptzM(M As CodeModule)
Dim Old$: Old = SrcLzM(M)
Dim NewLines$: NewLines = SSrcLzM(M)
Dim O$: O = IIf(Old = NewLines, "(Same)", "<====Diff")
Debug.Print Mdn(M), O
End Sub

Sub SrtMd()
SrtzM CMd
End Sub

Private Sub SrtzP(P As VBProject)
BackupFfn Pjf(P)
Dim C As VBComponent
For Each C In P.VBComponents
    SrtzM C.CodeModule
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
A1 = DclLy(A)
B1 = DclLy(B)
Stop
End Sub

Private Sub Z_SrtMd()
Dim Md As CodeModule
GoSub X0
Exit Sub
X0:
    Dim I
'    For Each I In MdAy(CPj)
        Set Md = I
        If Mdn(Md) = "Str_" Then
            GoSub Ass
        End If
'    Next
    Return
X1:
    Return
Ass:
    Debug.Print Mdn(Md); vbTab;
    Dim BefSrt$(), AftSrt$()
    BefSrt = Src(Md)
    AftSrt = SplitCrLf(SSrcLzM(Md))
    If JnCrLf(BefSrt) = JnCrLf(AftSrt) Then
        Debug.Print "Is Same of before and after sorting ......"
        Return
    End If
    If Si(AftSrt) <> 0 Then
        If LasEle(AftSrt) = "" Then
            Dim Pfx
            Pfx = Array("There is non-blank-line at end after sorting", "Md=[" & Mdn(Md) & "=====")
            BrwAy AddAyAp(Pfx, AftSrt)
            Stop
        End If
    End If
    Dim A$(), B$(), II
    A = MinusAy(BefSrt, AftSrt)
    B = MinusAy(AftSrt, BefSrt)
    Debug.Print
    If Si(A) = 0 And Si(B) = 0 Then Return
    If Si(AeEmpEle(A)) <> 0 Then
        Debug.Print "Si(A)=" & Si(A)
        BrwAy A
        Stop
    End If
    If Si(AeEmpEle(B)) <> 0 Then
        Debug.Print "Si(B)=" & Si(B)
        BrwAy B
        Stop
    End If
    Return
End Sub

