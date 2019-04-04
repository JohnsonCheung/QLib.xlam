Attribute VB_Name = "MVb_Cml"
Option Explicit
#Const PlaySav = 1
Public Const DocOfCml0$ = "It a String with Fst"
Private Function MthDotCmlGpAsetzVbe(A As Vbe, Optional WhStr$) As Aset
Set MthDotCmlGpAsetzVbe = AsetzAy(MthDotCmlGpAyzVbe(A, WhStr))
End Function

Private Sub Z_CmlAset()
CmlAset(NyzStr(SrcLineszPj(CurPj))).Srt.Brw
End Sub

Private Sub Z_ShfCml()
Dim L$, EptL$
Ept = "A"
L = "AABcDD"
EptL = "ABcDD"
GoSub Tst
Exit Sub
Tst:
    Act = ShfCml(L)
    If Act <> Ept Then Stop
    If EptL <> L Then Stop
    Return
End Sub

Function Cml0Ay(Nm) As String()
If Nm = "" Then Exit Function
#If PlaySav Then
If Not IsNm(Nm) Then Thw CSub, "Given Nm is not a name", "Nm", Nm
#End If
Dim J&, Cml$, C$, A%, IsUCas As Boolean, IsLowZ As Boolean
Cml = FstChr(Nm)
For J = 2 To Len(Nm)
    C = Mid(Nm, J, 1)
    A = Asc(C)
    IsLowZ = C = "z"
    IsUCas = IsAscUCas(A)
    Select Case True
    Case IsUCas:  PushNonBlankStr Cml0Ay, Cml
                  Cml = C
    Case IsLowZ:  PushNonBlankStr Cml0Ay, Cml
                  PushI Cml0Ay, "z"
                  Cml = ""
    Case Else:    Cml = Cml & C
    End Select
Next
PushNonBlankStr Cml0Ay, Cml
End Function

Function Cml1Ay(Nm) As String()
Dim M$(), A$()
Dim InCombine As Boolean, I, IsOneChr As Boolean, IsLowZ As Boolean
A = Cml0Ay(Nm)
For Each I In Itr(A)
    IsOneChr = Len(I) = 1
    IsLowZ = I = "z"
    Select Case True
    Case IsLowZ:                    PushNonBlankStr Cml1Ay, Jn(M): Erase M: PushI Cml1Ay, "z"
    Case IsOneChr And InCombine:    PushI M, I
    Case IsOneChr:                  PushI M, I: InCombine = True
    Case InCombine:                 PushI M, I: PushI Cml1Ay, Jn(M): Erase M
    Case Else:                      PushI Cml1Ay, I
    End Select
Next
If Si(M) > 0 Then PushI Cml1Ay, Jn(M)
End Function

Function CmlAset(Ny$()) As Aset
Set CmlAset = New Aset
Dim Nm
For Each Nm In Itr(Ny)
    CmlAset.PushAy CmlAy(Nm)
Next
End Function

Function CmlAy(Nm) As String()
Dim I
For Each I In Cml1Ay(Nm)
    PushI CmlAy, RmvLDashSfx(RmvDigSfx(I))
Next
End Function

Function CmlAyzNy(Ny) As String()
Dim L
For Each L In Itr(Ny)
    PushI CmlAyzNy, CmlAy(L)
Next
End Function

Function CmlGpAy(Nm) As String()
Dim M$(), Cml, O$()
For Each Cml In CmlAy(Nm)
    Debug.Print Cml; "<--CmlQGpAy"
    If IsBRKCml(Cml) Then
        PushNonBlankStr O, Jn(M)
        PushI O, Cml
        Erase M
    Else
        PushI M, Cml
    End If
Next
PushNonBlankStr O, Jn(M)
CmlGpAy = O
Dim I
For Each I In Itr(O)
    If FstChr(I) = "_" Then Stop
Next
End Function

Function CmlLin$(Nm)
CmlLin = Nm & " " & JnSpc(Cml1Ay(Nm))
End Function

Function CmlLy(Ny) As String()
Dim L
For Each L In Itr(Ny)
    PushI CmlLy, CmlLin(L)
Next
End Function

Function CmlQGpAy(Nm) As String()
Dim IsVerbQuoted As Boolean, CmlQGp, O$()
For Each CmlQGp In CmlGpAy(Nm)
    Debug.Print CmlQGp; "<-- CmlGpAy"
    Select Case True
    Case Not IsVerbQuoted And IsVerb(CmlQGp): PushI O, QuoteSq(CmlQGp): IsVerbQuoted = True
    Case IsBRKCml(CmlQGp):                    PushI O, QuoteBkt(CmlQGp)
    Case Else:                                PushI O, CmlQGp
    End Select
Next
CmlQGpAy = O
End Function

Function CmlSetzNy(Ny$()) As Aset
Dim O As New Aset, S
For Each S In Itr(Ny)
    O.PushAy CmlAy(S)
Next
Set CmlSetzNy = O
End Function

Function DotCml$(Nm)
DotCml = JnQDot(CmlAy(Nm))
End Function

Function DotCmlGp$(Nm) ' = JnQDot . CmpGpAy
DotCmlGp = JnQDot(CmlGpAy(Nm))
End Function

Function DotCmlQGp$(Nm) ' = JnQDot . CmpGp1Ay
Dim O$: O = JnQDot(CmlQGpAy(Nm))
DotCmlQGp = O
'If HasEle(Array(".z.EFSchm.", _
".z.FFFxw.", _
".z.NEmpty.", _
".z.NEmpty.", _
".z.RRCCJ.", _
".z.RRCCSq.", _
".z.SAy.", _
".z.SDotDTimFfn."), O) Then Debug.Print Nm, O
End Function

Function FstCml$(S)
FstCml = ShfCml(CStr(S))
End Function

Function FstCmlAy(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushI FstCmlAy, FstCml(I)
Next
End Function

Function FstCmlx$(S)
FstCmlx = S & " " & FstCml(S)
End Function

Function FstCmlxAy(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushI FstCmlxAy, FstCmlx(I)
Next
End Function

Function FstCmlzWithSng$(S)
Dim Lin$, A$, O$, J%
Lin = S
While Lin <> ""
    J = J + 1: If J > 1000 Then ThwLoopingTooMuch CSub
    A = ShfCml(Lin)
    Select Case Len(A)
    Case 1: O = O & A
    Case Else: FstCmlzWithSng = O & A: Exit Function
    End Select
Wend
FstCmlzWithSng = O
End Function

Function IsAscCmlChr(A%) As Boolean
Select Case True
Case IsAscLetter(A), IsAscDig(A), IsAscLDash(A): IsAscCmlChr = True
End Select
End Function

Function IsAscFstCmlChr(A%) As Boolean
If IsAscLDash(A) Then Exit Function
IsAscFstCmlChr = IsAscCmlChr(A)
End Function

Function IsBRKCml(Cml) As Boolean
Select Case True
Case BRKCmlASet.Has(Cml), Cml = "z", IsULCml(Cml): IsBRKCml = True
End Select
End Function

Function IsULCml(Cml) As Boolean
Select Case True
Case Len(Cml) <> 2, Not IsAscUCas(FstAsc(Cml)), Not IsAscLCas(SndAsc(Cml))
Case Else: IsULCml = LCase(FstChr(Cml)) = SndChr(Cml)
End Select
End Function

Function MthDotCmlGpAsetOfVbe(Optional WhStr$) As Aset
Set MthDotCmlGpAsetOfVbe = MthDotCmlGpAsetzVbe(CurVbe, WhStr)
End Function

Function MthDotCmlGpAyOfVbe(Optional WhStr$) As String()
MthDotCmlGpAyOfVbe = MthDotCmlGpAyzVbe(CurVbe, WhStr)
End Function

Function MthDotCmlGpAyzVbe(A As Vbe, Optional WhStr$) As String()
Dim MthNm
For Each MthNm In MthNyzVbe(A, WhStr)
    PushI MthDotCmlGpAyzVbe, DotCmlGp(MthNm)
Next
End Function

Function RmvDigSfx$(S)
Dim J%
For J = Len(S) To 1 Step -1
    If Not IsAscDig(Asc(Mid(S, J, 1))) Then RmvDigSfx = Left(S, J): Exit Function
Next
End Function

Function RmvLDashSfx$(S)
Dim J%
For J = Len(S) To 1 Step -1
    If Mid(S, J, 1) <> "_" Then RmvLDashSfx = Left(S, J): Exit Function
Next
End Function

Function Seg1ErNy() As String()
Erase XX
X "Act"
X "App"
X "Ass"
X "Ay"
X "Bar"
X "Brk"
X "C3"
X "C4"
X "Can"
X "Cell"
X "Cm"
X "Cmd"
X "Db"
X "Dbtt"
X "Dic"
X "Dry"
X "Ds"
X "Ent"
X "F"
X "Fb"
X "Fbq"
X "Fdr"
X "Fny"
X "Frm"
X "Fun"
X "Fx"
X "Git"
X "Has"
X "Lg"
X "Lgr"
X "Lnx"
X "Lo"
X "Md"
X "Min"
X "Msg"
X "Mth"
X "N"
X "O"
X "Pc"
X "Pj"
X "Ps1"
X "Pt"
X "Pth"
X "Re"
X "Res"
X "Rs"
X "Scl"
X "Sess"
X "Shp"
X "Spec"
X "Sql"
X "Sw"
X "T"
X "Tak"
X "Tim"
X "Tmp"
X "To"
X "Txtb"
X "V"
X "W"
X "Xls"
X "Y"
Seg1ErNy = XX
End Function

Function ShfCml$(OStr$)
Dim J&, Fst As Boolean, Cml$, C$, A%, IsNmChr As Boolean, IsFstNmChr As Boolean
Fst = True
For J = 1 To Len(OStr)
    C = Mid(OStr, J, 1)
    A = Asc(C)
    IsNmChr = IsAscNmChr(A)
    IsFstNmChr = IsAscFstNmChr(A)
    Select Case True
    Case Fst
        Cml = C
        Fst = False
    Case IsAscUCas(A)
        If Cml <> "" Then GoTo R
        Cml = C
    Case IsAscDig(A)
        If Cml <> "" Then Cml = Cml & C
    Case IsAscLCas(A)
        Cml = Cml & C
    Case Else
        If Cml <> "" Then GoTo R
        Cml = ""
    End Select
Next
R:
    ShfCml = Cml
    OStr = Mid(OStr, J)
End Function

Function ShfCmlAy(S) As String()
Dim L$: L = S
Dim J&
While True
    J = J + 1: If J > 1000 Then ThwLoopingTooMuch CSub
    PushNonBlankStr ShfCmlAy, ShfCml(L)
    If L = "" Then Exit Function
Wend
End Function

Sub VcMthDotCmlGpAsetOfVbe(Optional WhStr$)
MthDotCmlGpAsetOfVbe.Srt.Vc
End Sub

Sub Z_Cml1Ay()
Dim Ny$(): Ny = MthNyOfVbe
Dim N
For Each N In Ny
    If N <> Jn(Cml1Ay(N)) Then Stop
Next
End Sub

