Attribute VB_Name = "QVb_Dta_Cml"
Option Explicit
Private Const CMod$ = "MVb_Cml."
Private Const Asm$ = "QVb"
#Const PlaySav = 1
Public Const DoczCml0$ = "It a String with FstChr is UCase"
Public Const DoczCml1$ = "It a String with FstChr is UCase"
Public Const DoczCmlLin = "It a Lin of Cml separated by spc"
Private Function MthDotCmlGpAsetzV(A As Vbe, Optional WhStr$) As Aset
Set MthDotCmlGpAsetzV = AsetzAy(MthDotCmlGpzV(A, WhStr))
End Function

Private Sub Z_CmlAset()
CmlAset(NyzStr(SrcLineszP(CPj))).Srt.Brw
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

Function Cml0Sy(Nm) As String()
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
    Case IsUCas:  PushNonBlank Cml0Sy, Cml
                  Cml = C
    Case IsLowZ:  PushNonBlank Cml0Sy, Cml
                  PushI Cml0Sy, "z"
                  Cml = ""
    Case Else:    Cml = Cml & C
    End Select
Next
PushNonBlank Cml0Sy, Cml
End Function

Function Cml1Ay(Nm) As String()
Dim M$(), A$()
Dim InCombine As Boolean, I, IsOneChr As Boolean, IsLowZ As Boolean
A = Cml0Sy(Nm)
For Each I In Itr(A)
    IsOneChr = Len(I) = 1
    IsLowZ = I = "z"
    Select Case True
    Case IsLowZ:                    PushNonBlank Cml1Ay, Jn(M): Erase M: PushI Cml1Ay, "z"
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
Dim INm
For Each INm In Itr(Ny)
    CmlAset.PushAy CmlSy(CStr(INm))
Next
End Function

Function CmlSy(Nm) As String()
Dim I
For Each I In Cml1Ay(Nm)
    PushI CmlSy, RmvLDashSfx(RmvDigSfx(CStr(I)))
Next
End Function

Function CmlSyzNy(Ny$()) As String()
Dim I, Nm$
For Each I In Itr(Ny)
    Nm = I
    PushI CmlSyzNy, CmlSy(Nm)
Next
End Function

Function CmlGp(Nm) As String()
Dim M$(), I, Cml$, O$()
For Each I In CmlSy(Nm)
    Cml = I
    Debug.Print Cml; "<--CmlQBlk"
    If IsBRKCml(Cml) Then
        PushNonBlank O, Jn(M)
        PushI O, Cml
        Erase M
    Else
        PushI M, Cml
    End If
Next
PushNonBlank O, Jn(M)
CmlGp = O
For Each I In Itr(O)
    Cml = I
    If FstChr(Cml) = "_" Then Stop
Next
End Function

Function CmlLin(Nm)
CmlLin = Nm & " " & JnSpc(Cml1Ay(Nm))
End Function

Function CmlLy(Ny$()) As String()
Dim L
For Each L In Itr(Ny)
    PushI CmlLy, CmlLin(CStr(L))
Next
End Function

Function CmlQBlk(Nm) As String()
Dim IsVerbQuoted As Boolean, CmlQGp$, I, O$()
For Each I In CmlGp(Nm)
    CmlQGp = I
    Debug.Print CmlQGp; "<-- CmlGp"
    Select Case True
    Case Not IsVerbQuoted And IsVerb(CmlQGp): PushI O, QuoteSq(CmlQGp): IsVerbQuoted = True
    Case IsBRKCml(CmlQGp):                    PushI O, QuoteBkt(CmlQGp)
    Case Else:                                PushI O, CmlQGp
    End Select
Next
CmlQBlk = O
End Function

Function CmlSetzNy(Ny$()) As Aset
Dim O As New Aset, I
For Each I In Itr(Ny)
    O.PushAy CmlSy(CStr(I))
Next
Set CmlSetzNy = O
End Function

Function DotCml$(Nm)
DotCml = JnQDot(CmlSy(Nm))
End Function

Function DotCmlGp$(Nm) ' = JnQDot . CmpBlk
DotCmlGp = JnQDot(CmlGp(Nm))
End Function

Function DotCmlQGp$(Nm) ' = JnQDot . CmpGp1Ay
Dim O$: O = JnQDot(CmlQBlk(Nm))
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

Function FstCmlSy(Sy$()) As String()
Dim I
For Each I In Itr(Sy)
    PushI FstCmlSy, FstCml(CStr(I))
Next
End Function

Function FstCmlx$(S)
FstCmlx = S & " " & FstCml(S)
End Function

Function FstCmlxzSy(Sy$()) As String()
Dim I
For Each I In Itr(Sy)
    PushI FstCmlxzSy, FstCmlx(CStr(I))
Next
End Function

Function FstCmlzWiSng$(S)
Dim Lin, A$, O$, J%
Lin = S
While Lin <> ""
    J = J + 1: If J > 1000 Then ThwLoopingTooMuch CSub
    A = ShfCml(Lin)
    Select Case Len(A)
    Case 1: O = O & A
    Case Else: FstCmlzWiSng = O & A: Exit Function
    End Select
Wend
FstCmlzWiSng = O
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

Function IsBRKCml(Cml$) As Boolean
Select Case True
Case BRKCmlASet.Has(Cml), Cml = "z", IsULCml(Cml): IsBRKCml = True
End Select
End Function

Function IsULCml(Cml$) As Boolean
Select Case True
Case Len(Cml) <> 2, Not IsAscUCas(FstAsc(Cml)), Not IsAscLCas(SndAsc(Cml))
Case Else: IsULCml = LCase(FstChr(Cml)) = SndChr(Cml)
End Select
End Function

Function MthDotCmlGpAsetInVbe(Optional WhStr$) As Aset
Set MthDotCmlGpAsetInVbe = MthDotCmlGpAsetzV(CVbe, WhStr)
End Function

Function MthDotCmlGpInVbe(Optional WhStr$) As String()
MthDotCmlGpInVbe = MthDotCmlGpzV(CVbe, WhStr)
End Function

Function MthDotCmlGpzV(A As Vbe, Optional WhStr$) As String()
Dim Mthn
For Each Mthn In MthNyzV(A, WhStr)
    PushI MthDotCmlGpzV, DotCmlGp(CStr(Mthn))
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

Function ShfCml$(OStr)
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

Function ShfCmlSy(S) As String()
Dim L$: L = S
Dim J&
While True
    J = J + 1: If J > 1000 Then ThwLoopingTooMuch CSub
    PushNonBlank ShfCmlSy, ShfCml(L)
    If L = "" Then Exit Function
Wend
End Function

Sub VcMthDotCmlGpAsetInVbe(Optional WhStr$)
MthDotCmlGpAsetInVbe.Srt.Vc
End Sub

Private Sub Z_Cml1Ay()
Dim Ny$(): Ny = MthnyV
Dim N
For Each N In Ny
    If N <> Jn(Cml1Ay(CStr(N))) Then Stop
Next
End Sub

