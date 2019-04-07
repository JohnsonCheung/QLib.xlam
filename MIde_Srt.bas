Attribute VB_Name = "MIde_Srt"
Option Explicit


Private Sub SrtzPj(A As VBProject)
FfnBackup Pjf(A)
RplPj A, SrtedMdDiczPj(A)
End Sub

Private Sub ZZ()
End Sub

Private Sub ZZ_Dcl_BefAndAft_Srt()
Const MdNm$ = "VbStrRe"
Dim A$() ' Src
Dim B$() ' Src->Srt
Dim A1$() 'Src->Dcl
Dim B1$() 'Src->Src->Dcl
A = Src(Md(MdNm))
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
'    For Each I In MdAy(CurPj)
        Set Md = I
        If MdNm(Md) = "Str_" Then
            GoSub Ass
        End If
'    Next
    Return
X1:
    Return
Ass:
    Debug.Print MdNm(Md); vbTab;
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
            Pfx = Array("There is non-blank-line at end after sorting", "Md=[" & MdNm(Md) & "=====")
            BrwAy AyAddAp(Pfx, AftSrt)
            Stop
        End If
    End If
    Dim A$(), B$(), II
    A = AyMinus(BefSrt, AftSrt)
    B = AyMinus(AftSrt, BefSrt)
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

Private Sub ZZ_SrtedSrcLineszMd()
BrwStr SrtedSrcLineszMd(Md("MIde_Md"))
End Sub

Private Sub Z_SrcLinesSrt()
Brw SrcLinesSrt(CurSrc)
End Sub

Function MthNm3zDNm(MthDNm) As MthNm3
Dim Nm$, Ty$, Mdy$
If MthDNm = "*Dcl" Then
    Nm = "*Dcl"
Else
    Dim B$(): B = SplitDot(MthDNm)
    If Si(B) <> 3 Then
        Thw CSub, "Given MthDNm SplitDot should be 3 elements", "NEle-SplitDot MthDNm", Si(B), MthDNm
    End If
    Dim ShtMdy$, ShtTy$
    AsgAp B, Nm, ShtTy, ShtMdy
    Ty = MthTyBySht(ShtTy)
    Mdy = ShtMthMdy(ShtMdy)
End If
Set MthNm3zDNm = New MthNm3
With MthNm3zDNm
    .Nm = Nm
    .MthMdy = Mdy
    .MthTy = Ty
End With
End Function

Function MthSrtKey$(MthDNm)
If MthDNm = "*Dcl" Then MthSrtKey = "*Dcl": Exit Function
Dim A$(): A = SplitDot(MthDNm): If Si(A) <> 3 Then Thw CSub, "Invalid MthDNm, should have 2 dot", "MthDNm", MthDNm
MthSrtKey = A(2) & "." & A(1) & "." & A(0)
End Function

Function SrtedSrczMd(A As CodeModule) As String()
SrtedSrczMd = SrcSrt(Src(A))
End Function
Function SrtedMdDicOfPj() As Dictionary
Set SrtedMdDicOfPj = SrtedMdDiczPj(CurPj)
End Function

Function SrtedMdDiczPj(A As VBProject) As Dictionary
Dim C As VBComponent, O As New Dictionary
For Each C In A.VBComponents
    O.Add C.Name, SrtedSrcLineszMd(C.CodeModule)
Next
Set SrtedMdDiczPj = O
End Function

Function SrcSrt(Src$()) As String()
SrcSrt = SplitCrLf(SrcLinesSrt(Src))
End Function

Function SrcDic(Src$(), Optional WiTopRmk As Boolean) As Dictionary
Dim D As Dictionary: Set D = MthDic(Src, WiTopRmk)
Dim K
Set SrcDic = New Dictionary
For Each K In D
    SrcDic.Add MthSrtKey(K), D(K)
Next
End Function

Function SrtedSrcDic(Src$()) As Dictionary
Set SrtedSrcDic = DicSrt(SrcDic(Src))
End Function

Function SrcLinesSrt$(Src$())
SrcLinesSrt = JnDblCrLf(SrtedSrcDic(Src).Items)
End Function

Function SrtedSrcLinesOfMd$()
SrtedSrcLinesOfMd = SrtedSrcLineszMd(CurMd)
End Function
Function SrtedSrcLineszMd$(A As CodeModule)
SrtedSrcLineszMd = SrcLinesSrt(Src(A))
End Function

Sub BrwSrtRptzMd(A As CodeModule)
Dim Old$: Old = SrcLineszMd(A)
Dim NewLines$: NewLines = SrtedSrcLineszMd(A)
Dim O$: O = IIf(Old = NewLines, "(Same)", "<====Diff")
Debug.Print MdNm(A), O
End Sub

Sub BrwSrtedMdDic()
BrwDic SrtedMdDicOfPj(CurPj)
End Sub

Sub RplPj(A As VBProject, MdDic As Dictionary)
Dim MdNm
For Each MdNm In MdDic.Keys
    If MdNm <> "MIde_Srt" Then
        RplMd MdzPj(A, MdNm), MdDic(MdNm)
    End If
Next
End Sub

Sub Srt()
SrtzMd CurMd
End Sub

Sub SrtPj()
SrtzPj CurPj
End Sub

Sub SrtzMd(A As CodeModule)
Dim Nm$: Nm = MdNm(A)
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
