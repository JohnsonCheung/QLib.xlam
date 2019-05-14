Attribute VB_Name = "QIde_Mth_MthDic"
Option Explicit
Private Const CMod$ = "MIde_Mth_Dic."
Private Const Asm$ = "QIde"
Function MthDiczP(P As VBProject) As Dictionary
Dim C As VBComponent
For Each C In P.VBComponents
    PushDic MthDiczP, MthDic(C.CodeModule)
Next
End Function

Private Sub ZZ_SMthDicM()
B SMthDicM
End Sub
Function SMthDicM() As Dictionary
Set SMthDicM = SMthDIczM(CMd)
End Function
Function SMthDIczM(A As CodeModule) As Dictionary
Set SMthDIczM = SrtDic(MthDiczM(A))
End Function

Private Sub Z_PjMthDic()
Dim A As Dictionary, V, K
Set A = Pj_MthDic(CPj)
Ass IsDiczSy(A) '
For Each K In A
    If InStr(K, ".") > 0 Then Stop
    If Si(A(K)) = 0 Then Stop
Next
End Sub

Private Sub Z_PjMthDic1()
Dim A As Dictionary, V, K
Set A = Pj_MthDic(CPj)
Ass IsDiczSy(A) '
For Each K In A
    If InStr(K, ".") > 0 Then Stop
    If Si(A(K)) = 0 Then Stop
Next
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

Function MthDicByDicDic(MthDicWiTopRmk As Dictionary, TopRmkDic As Dictionary) As Dictionary
Dim K, O As New Dictionary
For Each K In MthDicWiTopRmk.Keys
    If TopRmkDic.Exists(K) Then
        O.Add K, TopRmkDic(K) & vbCrLf & MthDicWiTopRmk(K)
    Else
        O.Add K, MthDicWiTopRmk(K)
    End If
Next
Set MthDicByDicDic = O
End Function

Function MthDicM() As Dictionary
Set MthDicM = MthDic(CMd)
End Function

Private Function MthDnzLin$(Lin)
MthDnzLin = MthDnzMthn3(Mthn3zLin(Lin))
End Function

Function MthDiczSN(Src$(), Mdn) As Dictionary 'Key is MthDn, Val is MthLinesWiTopRmk
Dim Ix, Lines$, Dn$
Set MthDiczSN = New Dictionary
Dim P$: P = Mdn & "."
With MthDiczSN
    .Add P & "*Dcl", Dcl(Src)
    For Each Ix In MthIxItr(Src)
        Dn = MthDnzLin(Src(Ix))
        Lines = MthLineszSI(Src, Ix, WiTopRmk:=True)
        .Add P & Dn, Lines
    Next
End With
End Function

Function MthDic(A As CodeModule) As Dictionary
Set MthDic = MthDiczSN(Src(A), Mdn(A))
End Function

Function LineszJnLinesItr$(LinesItr, Optional Sep$ = vbCrLf)
LineszJnLinesItr = Jn(IntozItr(EmpSy, LinesItr), Sep)
End Function

Function SMthDicP() As Dictionary
Set SMthDicP = SMthDiczP(CPj)
End Function

Function SMthDiczP1(P As VBProject) As Dictionary
Set SMthDiczP1 = SrtDic(MthDiczP(P))
End Function

Function SSrcDicM() As Dictionary
Set SSrcDicM = SSrcDic(CSrc)
End Function

Function SSrcDic(Src$(), Mdn) As Dictionary
Set SSrcDic = SrtDic(SrcDic(Src, Mdn))
End Function

Function SSrcDiczM$(A As CodeModule)
SrtedSrcLineszMd = JnCrLf(SrcSrt(Src(A)))
End Function

Sub BrwSrtRptzMd(A As CodeModule)
Dim Old$: Old = SrcLineszMd(A)
Dim NewLines$: NewLines = SrtedSrcLineszMd(A)
Dim O$: O = IIf(Old = NewLines, "(Same)", "<====Diff")
Debug.Print Mdn(A), O
End Sub

Sub BrwSrtedMdDic()
BrwDic SrtedMdDicP(CPj)
End Sub

Sub RplPj(P As VBProject, MdDic As Dictionary)
Dim Mdn
For Each Mdn In MdDic.Keys
    If Mdn <> "MIde_Srt" Then
        RplMd MdzPN(P, Mdn), MdDic(Mdn)
    End If
Next
End Sub

Sub SrtP()
SrtzP CPj
End Sub
Sub SrtM()
SrtMd CMd
End Sub
Sub SrtMd()
SrtMdzM CMd
End Sub
Sub SrtMdzM(A As CodeModule)
RplMd A, SrtedSrcLineszMd(A)
Exit Sub
Dim Nm$: Nm = Mdn(A)
Debug.Print "Sorting: "; AlignL(Nm, 20); " ";
Dim LinesN$: LinesN = SrtedSrcLineszMd(A)
Dim LinesO$: LinesO = SrcLineszMd(A)
'Exit if same
    If LinesO = LinesN Then
        Debug.Print "<== Same"
        Exit Sub
    End If
'Delete
    Debug.Print FmtQQ("<--- Deleted (?) lines", A.CountOfLines);
    ClrMd A
'Add sorted lines
    A.AddFromString LinesN
    Debug.Print "<----Sorted Lines added...."
End Sub

Private Sub SrtzP(P As VBProject)
BackupFfn Pjf(P)
RplPj P, MdygPOfSrt(P)
End Sub

Private Sub ZZ_Dcl_BefAndAft_Srt()
Const Mdn = "VbStrRe"
Dim A$() ' Src
Dim B$() ' Src->Srt
Dim A1$() 'Src->Dcl
Dim B1$() 'Src->Src->Dcl
A = Src(Md(Mdn))
B = SrcSrt(A)
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
    AftSrt = SplitCrLf(SrtedSrcLineszMd(Md))
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

