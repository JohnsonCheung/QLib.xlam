Attribute VB_Name = "QIde_Mth_MthDic"
Option Compare Text
Option Explicit
Const CMod$ = "MIde_Mth_Dic."
Const Asm$ = "QIde"
Public Const DoczSMdDic$ = "SMdDic is Sorted-Mdn-To-SrcLines."

Function MthDiczP(P As VBProject) As Dictionary
Dim C As VBComponent
For Each C In P.VBComponents
    PushDic MthDiczP, MthDiczM(C.CodeModule)
Next
End Function

Private Sub ZZ_SMthDicM()
B SMthDicM
End Sub
Function SMthDicM() As Dictionary
Set SMthDicM = SMthDIczM(CMd)
End Function
Function SMthDIczM(M As CodeModule) As Dictionary
Set SMthDIczM = SrtDic(MthDiczM(M))
End Function

Private Sub ZZ_MthDiczP()
Dim A As Dictionary: Set A = MthDiczP(CPj)
Ass IsDicOfLines(A) '
Vc A
End Sub

Private Sub ZZ_MthDicM()
B MthDicM
End Sub

Function MthDicP()
Set MthDicP = MthDiczP(CPj)
End Function

Function CSMthDicP() As Dictionary
Static X As Boolean, Y As Dictionary
If Not X Then
    X = True
    Set Y = SMthDicP
End If
Set CSMthDicP = Y
End Function

Function MthDicM() As Dictionary
Set MthDicM = MthDiczM(CMd)
End Function

Private Function MthDnzLin$(Lin)
MthDnzLin = MthDnzMthn3(Mthn3zL(Lin))
End Function

Function MthDic(Src$(), Optional Mdn$, Optional ExlDcl As Boolean) As Dictionary 'Key is MthDn, Val is MthLinesWiTopRmk
Dim Ix, Lines$, Dn$
Set MthDic = New Dictionary
Dim P$: If Mdn <> "" Then P = Mdn & "."
With MthDic
    If Not ExlDcl Then .Add P & "*Dcl", Dcl(Src)
    For Each Ix In MthIxItr(Src)
        Dn = MthDnzLin(Src(Ix))
        Lines = MthLineszSI(Src, Ix)
        .Add P & Dn, Lines
    Next
End With
End Function

Function MthDiczM(M As CodeModule) As Dictionary
Set MthDiczM = MthDic(Src(M), Mdn(M))
End Function

Function LineszJnLinesItr$(LinesItr, Optional Sep$ = vbCrLf)
LineszJnLinesItr = Jn(IntozItr(EmpSy, LinesItr), Sep)
End Function
Function SMthDiczP(P As VBProject) As Dictionary
Set SMthDiczP = SrtDic(MthDiczP(P))
End Function
Function SMthDicP() As Dictionary
Set SMthDicP = SMthDiczP(CPj)
End Function

Function SSrcLineszS$(Src$())
SSrcLineszS = JnStrDic(SrtDic(MthDic(Src)), vbDblCrLf)
End Function

Function SSrcLineszM$(M As CodeModule)
SSrcLineszM = SSrcLineszS(Src(M))
End Function

Function SrcLineszM$(M As CodeModule)
If M.CountOfLines > 0 Then
    SrcLineszM = M.Lines(1, M.CountOfLines)
End If
End Function

Sub BrwSrtRptzM(M As CodeModule)
Dim Old$: Old = SrcLineszM(M)
Dim NewLines$: NewLines = SSrcLineszM(M)
Dim O$: O = IIf(Old = NewLines, "(Same)", "<====Diff")
Debug.Print Mdn(M), O
End Sub

Sub BrwSMdDiczP()
BrwDic SMdDiczP(CPj)
End Sub

Sub RplPj(A As RplgPj)
Dim Mdn, J&
For J = 0 To A.RplgMds.N - 1
    If Mdn <> "MIde_Srt" Then
        RplMd A.RplgMds.Ay(J)
    End If
Next
End Sub

Sub SrtP()
SrtzP CPj
End Sub
Sub SrtM()
SrtzM CMd
End Sub
Sub SrtMd()
SrtzM CMd
End Sub

Sub SrtzM(M As CodeModule)
With SomSrtgzM(M)
    If .Som Then RplMd .RplgMd
End With
End Sub

Private Sub SrtzP(P As VBProject)
BackupFfn Pjf(P)
RplPj SrtgzP(P)
End Sub

Function SrtgzP(P As VBProject) As RplgPj
Dim O As RplgPj, M As CodeModule
Set O.Pj = P
Dim C As VBComponent
For Each C In P.VBComponents
    Set M = C.CodeModule
    With SomSrtgzM(M)
        If .Som Then
            PushRplgMd O.RplgMds, .RplgMd
        End If
    End With
Next
End Function

Function SomSrtgzM(M As CodeModule) As SomRplgMd
Dim Srted$
    Dim S$(): S = Src(M)
    Srted = SSrcLineszS(S)
Dim NotSrted$
    NotSrted = JnCrLf(S)
If Srted = NotSrted Then Exit Function
SomSrtgzM = SomRplgMd(RplgMd(M, Srted))
End Function

Function SomRplgMd(RplgMd As RplgMd) As SomRplgMd
With SomRplgMd
    .Som = True
    .RplgMd = RplgMd
End With
End Function

Private Sub ZZ_Dcl_BefAndAft_Srt()
Const Mdn = "VbStrRe"
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

Private Sub ZZ_SrtMd()
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
    AftSrt = SplitCrLf(SSrcLineszM(Md))
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
    If Si(AyeEmpEle(A)) <> 0 Then
        Debug.Print "Si(A)=" & Si(A)
        BrwAy A
        Stop
    End If
    If Si(AyeEmpEle(B)) <> 0 Then
        Debug.Print "Si(B)=" & Si(B)
        BrwAy B
        Stop
    End If
    Return
End Sub

