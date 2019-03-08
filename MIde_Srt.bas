Attribute VB_Name = "MIde_Srt"
Option Explicit
Function SrtedMdDic(A As VBProject) As Dictionary
Dim C As VBComponent, O As New Dictionary
For Each C In A.VBComponents
    O.Add C.Name, SrtedSrcLineszMd(C.CodeModule)
Next
Set SrtedMdDic = O
End Function
Sub BrwSrtedMdDic()
BrwDic SrtedMdDic(CurPj)
End Sub

Function SrtedSrcLines$(Src$())
SrtedSrcLines = JnDblCrLf(SrtedSrcDic(Src).Items)
End Function

Sub BrwSrtRptzMd(A As CodeModule)
Dim Old$: Old = LinesMd(A)
Dim NewLines$: NewLines = SrtedSrcLineszMd(A)
Dim O$: O = IIf(Old = NewLines, "(Same)", "<====Diff")
Debug.Print MdNm(A), O
End Sub
Sub Srt()
SrtMd CurMd
End Sub

Sub SrtMd(A As CodeModule)
Dim Nm$: Nm = MdNm(A)
Debug.Print "Sorting: "; AlignL(Nm, 20); " ";
Dim LinesN$: LinesN = SrtedSrcLineszMd(A)
Dim LinesO$: LinesO = LinesMd(A)
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

Function SrtedSrcLineszMd$(A As CodeModule)
SrtedSrcLineszMd = SrtedSrcLines(Src(A))
End Function

Function SrtedMd(A As CodeModule) As String()
SrtedMd = SrtedSrc(Src(A))
End Function
Function MthNm3zDNm(MthDNm) As MthNm3
Dim Nm$, Ty$, Mdy$
If MthDNm = "*Dcl" Then
    Nm = "*Dcl"
Else
    Dim B$(): B = SplitDot(MthDNm)
    If Sz(B) <> 3 Then
        Thw CSub, "Given MthDNm SplitDot should be 3 elements", "NEle-SplitDot MthDNm", Sz(B), MthDNm
    End If
    Dim ShtMdy$, ShtTy$
    AsgAp B, Nm, ShtTy, ShtMdy
    Ty = MthTySht(ShtTy)
    Mdy = MthMdySht(ShtMdy)
End If
Set MthNm3zDNm = New MthNm3
With MthNm3zDNm
    .Nm = Nm
    .MthMdy = Mdy
    .MthTy = Ty
End With
End Function

Sub SrtPj(A As VBProject)
Dim MdNm, D As Dictionary
Set D = SrtedMdDic(A)
For Each MdNm In D.Keys
    If MdNm <> "MIde_Srt" Then
        RplMd MdzPj(A, MdNm), D(MdNm)
    End If
Next
End Sub

Function SrtedSrcDic(Src$()) As Dictionary
Set SrtedSrcDic = DicSrt(MthDic(Src))
End Function

Function SrtedSrc(Src$()) As String()
SrtedSrc = SplitCrLf(SrtedSrcLines(Src))
End Function

Private Sub ZZ_Dcl_BefAndAft_Srt()
Const MdNm$ = "VbStrRe"
Dim A$() ' Src
Dim B$() ' Src->Srt
Dim A1$() 'Src->Dcl
Dim B1$() 'Src->Src->Dcl
A = Src(Md(MdNm))
B = SrtedSrc(A)
A1 = DclLy(A)
B1 = DclLy(B)
Stop
End Sub

Private Sub Z_SrtedSrcLines()
Brw SrtedSrcLines(SrcMd)
End Sub

Private Sub ZZ()
End Sub

Private Sub ZZ_SrtedSrcLineszMd()
BrwStr SrtedSrcLineszMd(Md("MIde_Md"))
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
    If Sz(AftSrt) <> 0 Then
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
    If Sz(A) = 0 And Sz(B) = 0 Then Return
    If Sz(AyeEmpEle(A)) <> 0 Then
        Debug.Print "Sz(A)=" & Sz(A)
        BrwAy A
        Stop
    End If
    If Sz(AyeEmpEle(B)) <> 0 Then
        Debug.Print "Sz(B)=" & Sz(B)
        BrwAy B
        Stop
    End If
    Return
End Sub



