Attribute VB_Name = "MIde_Srt"
Option Explicit
Function SrtedMdDic(A As VBProject) As Dictionary
Dim C As VBComponent, O As New Dictionary
For Each C In A.VBComponents
    O.Add C.Name, SrtedSrcLinzMd(C.CodeModule)
Next
Set SrtedMdDic = O
End Function
Sub BrwSrtedMdDic()
BrwDic SrtedMdDic(CurPj)
End Sub
Function SrtedSrcLinesz$(Src$())
SrtedSrcLinesz = JnDblCrLf(SrtedSrcDic(Src).Items)
End Function

Property Get SrtedSrcLines$()
SrtedSrcLines = SrtedSrcLinzMd(CurMd)
End Property

Function MthSrtKey$(Lin)
MthSrtKey = MthSrtKeyzNm3(MthNm3(Lin))
End Function

Sub BrwSrtRptzMd(A As CodeModule)
Dim Old$: Old = LinesMd(A)
Dim NewLines$: NewLines = SrtedSrcLinzMd(A)
Dim O$: O = IIf(Old = NewLines, "(Same)", "<====Diff")
Debug.Print MdNm(A), O
End Sub
Sub Srt()
SrtMd CurMd
End Sub

Sub SrtMd(A As CodeModule)
Dim Nm$: Nm = MdNm(A)
Debug.Print "Sorting: "; AlignL(Nm, 20); " ";
Dim LinesN$: LinesN = SrtedSrcLinzMd(A)
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

Function SrtedSrcLinzMd$(A As CodeModule)
SrtedSrcLinzMd = SrtedSrcLinesz(Src(A))
End Function
Property Get SrtedSrc() As String()
SrtedSrc = SrtedSrczMd(CurMd)
End Property
Function SrtedSrczMd(A As CodeModule) As String()
SrtedSrczMd = SrtedSrcz(Src(A))
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

Function MthSrtKeyzDNm$(MthDNm) ' MthDNm is Nm.Ty.Mdy
MthSrtKeyzDNm = MthSrtKeyzNm3(MthNm3zDNm(MthDNm))
End Function

Function MthSrcKeyAyzDNm(MthDNy$()) As String() ' MthDNm is Nm.Ty.Mdy
Dim I
For Each I In Itr(MthDNy)
    PushI MthSrcKeyAyzDNm, MthSrtKey(I)
Next
End Function

Function MthSrtKeyzNm3$(A As MthNm3)
Dim Mdy$, Ty$, Nm$
With A
    Nm = .Nm
    Ty = .ShtTy
    Mdy = .ShtMdy
End With
Dim P% 'Priority
    Select Case True
    Case HasPfx(Nm, "Init"): P = 1
    Case Nm = "Z":           P = 9
    Case Nm = "ZZ":          P = 8
    Case HasPfx(Nm, "Z_"):   P = 7
    Case HasPfx(Nm, "ZZ_"):  P = 6
    Case HasPfx(Nm, "Z"):    P = 5
    Case Else:               P = 2
    End Select
MthSrtKeyzNm3 = P & ":" & Nm & ":" & Ty & ":" & Mdy
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
Dim K, D As Dictionary, O As New Dictionary
Set D = MthDic(Src)
For Each K In D.Keys
    O.Add MthSrtKeyzDNm(K), D(K)
Next
Set SrtedSrcDic = O
End Function

Function SrtedSrcz(Src$()) As String()
SrtedSrcz = SplitCrLf(SrtedSrcLinesz(Src))
End Function

Private Sub ZZ_Dcl_BefAndAft_Srt()
Const MdNm$ = "VbStrRe"
Dim A$() ' Src
Dim B$() ' Src->Srt
Dim A1$() 'Src->Dcl
Dim B1$() 'Src->Src->Dcl
A = Src(Md(MdNm))
B = SrtedSrcz(A)
A1 = DclLy(A)
B1 = DclLy(B)
Stop
End Sub

Private Sub Z_MthSrtKey()
GoTo ZZ
Dim A$
'
Ept = "2:MthSrtKey_Lin:Function:": A = "Function MthSrtKey_Lin$(A)": GoSub Tst
Ept = "2:YYA:Function:":           A = "Function YYA()":            GoSub Tst
Exit Sub
Tst:
    Act = MthSrtKey(A)
    C
    Return
ZZ:
    Dim Ay1$(): Ay1 = MthLinAyzSrc(SrczVbe(CurVbe))
    Dim Ay2$(): Ay2 = MthSrtKeyAy(Ay1)
    Stop
    BrwS1S2Ay S1S2AyAyab(Ay2, Ay1)
End Sub

Private Sub Z_MthSrtKeyzDNm()
GoSub T1
'GoSub T2
Exit Sub
T1:

    Dim Ay1$(): Ay1 = MthDNyzSrc(SrcMd)
    Dim Ay2$(): Ay2 = MthSrcKeyAyzDNm(Ay1)
    BrwS1S2Ay S1S2AyAyab(Ay2, Ay1)
    Return
T2:
    Const A$ = "YYA.Fun."
    Debug.Print MthSrtKeyzDNm(A)
    Return
End Sub

Private Sub Z_SrtedSrcLinesz()
Brw SrtedSrcLinesz(SrcMd)
End Sub

Private Sub ZZ()
End Sub

Private Sub Z()
Z_MthSrtKey
End Sub

Private Sub ZZ_SrtedSrcLinzMd()
BrwStr SrtedSrcLinzMd(Md("MIde_Md"))
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
    AftSrt = SplitCrLf(SrtedSrcLinzMd(Md))
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



