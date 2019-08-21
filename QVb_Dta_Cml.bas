Attribute VB_Name = "QVb_Dta_Cml"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Cml."
Private Const Asm$ = "QVb"
':CmlAy: :
Private Function MthDotCmlGpAsetzV(A As Vbe) As Aset
Set MthDotCmlGpAsetzV = AsetzAy(MthDotCmlGpzV(A))
End Function

Private Sub Z_CmlAset()
CmlAset(NyzStr(SrcLzP(CPj))).Srt.Brw
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

Function CmlAy(Nm) As String()
'Ret : :Cml-ay.  Cml-ay is fm :Nm.  Each cml start with UCas and rest is LCas|Dig|_, ept fst cml the start letter may be LCas.

If Nm = "" Then Exit Function
#If PlaySav Then
If Not IsNm(Nm) Then Thw CSub, "Given Nm is not a name", "Nm", Nm
#End If
Dim J&, Cml$, C$, A%, O$()
Cml = FstChr(Nm)
For J = 2 To Len(Nm)
    C = Mid(Nm, J, 1)
    A = Asc(C)
    If IsAscUCas(A) Then
        PushNB O, Cml
        Cml = C
    Else
        Cml = Cml & C
    End If
Next
PushNB O, Cml
CmlAy = O
End Function

Function CmlAset(Ny$()) As Aset
Set CmlAset = New Aset
Dim INm
For Each INm In Itr(Ny)
    CmlAset.PushAy CmlAy(CStr(INm))
Next
End Function

Function CmlAyzNy(Ny$()) As String()
Dim I, Nm$
For Each I In Itr(Ny)
    Nm = I
    PushI CmlAyzNy, CmlAy(Nm)
Next
End Function

Function CmlGp(Nm) As String()
Dim M$(), I, Cml$, O$()
For Each I In CmlAy(Nm)
    Cml = I
    Debug.Print Cml; "<--CmlQBlk"
    If IsCmlBRK(Cml) Then
        PushNB O, Jn(M)
        PushI O, Cml
        Erase M
    Else
        PushI M, Cml
    End If
Next
PushNB O, Jn(M)
CmlGp = O
For Each I In Itr(O)
    Cml = I
    If FstChr(Cml) = "_" Then Stop
Next
End Function

Function Cmlss(Nm)
':Cmlss: :SS
Cmlss = Nm & " " & JnSpc(CmlAy(Nm))
End Function

Function CmlssAy(Ny$()) As String()
Dim L
For Each L In Itr(Ny)
    PushI CmlssAy, Cmlss(CStr(L))
Next
End Function

Function CmlQBlk(Nm) As String()
Dim IsVerbQted As Boolean, CmlQGp$, I, O$()
For Each I In CmlGp(Nm)
    CmlQGp = I
    Debug.Print CmlQGp; "<-- CmlGp"
    Select Case True
    Case Not IsVerbQted And IsVerb(CmlQGp): PushI O, QteSq(CmlQGp): IsVerbQted = True
    Case IsCmlBRK(CmlQGp):                    PushI O, QteBkt(CmlQGp)
    Case Else:                                PushI O, CmlQGp
    End Select
Next
CmlQBlk = O
End Function

Function CmlSetzNy(Ny$()) As Aset
Dim O As New Aset, I
For Each I In Itr(Ny)
    O.PushAy CmlAy(CStr(I))
Next
Set CmlSetzNy = O
End Function

Function DotCml$(Nm)
DotCml = QteJnDot(CmlAy(Nm))
End Function

Function DotCmlGp$(Nm) ' = QteJnDot . CmpBlk
DotCmlGp = QteJnDot(CmlGp(Nm))
End Function

Function DotCmlQGp$(Nm) ' = QteJnDot . CmpGp1Ay
Dim O$: O = QteJnDot(CmlQBlk(Nm))
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

Function FstCmlAy(Sy$()) As String()
Dim I
For Each I In Itr(Sy)
    PushI FstCmlAy, FstCml(CStr(I))
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

Function FstCmlzSng$(S)
Dim Lin, A$, O$, J%
Lin = S
While Lin <> ""
    J = J + 1: If J > 1000 Then ThwLoopingTooMuch CSub
    A = ShfCml(Lin)
    Select Case Len(A)
    Case 1: O = O & A
    Case Else: FstCmlzSng = O & A: Exit Function
    End Select
Wend
FstCmlzSng = O
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

Function IsCmlBRK(Cml$) As Boolean
Select Case True
Case BRKCmlASet.Has(Cml), Cml = "z", IsCmlUL(Cml): IsCmlBRK = True
End Select
End Function

Function IsCmlUL(Cml$) As Boolean
Select Case True
Case Len(Cml) <> 2, Not IsAscUCas(FstAsc(Cml)), Not IsAscLCas(SndAsc(Cml))
Case Else: IsCmlUL = LCase(FstChr(Cml)) = SndChr(Cml)
End Select
End Function

Function MthDotCmlGpAsetInVbe() As Aset
Set MthDotCmlGpAsetInVbe = MthDotCmlGpAsetzV(CVbe)
End Function

Function MthDotCmlGpInVbe() As String()
MthDotCmlGpInVbe = MthDotCmlGpzV(CVbe)
End Function

Function MthDotCmlGpzV(A As Vbe) As String()
Dim Mthn
For Each Mthn In MthNyzV(A)
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
X "Dy"
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
X "Exl"
X "Y"
Seg1ErNy = XX
End Function

Function ShfCml$(OStr)
Dim J&, Fst As Boolean, Cml$, C$, A%, IsChrzNm As Boolean, IsFstNmChr As Boolean
Fst = True
For J = 1 To Len(OStr)
    C = Mid(OStr, J, 1)
    A = Asc(C)
    IsChrzNm = IsAscNmChr(A)
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
    PushNB ShfCmlAy, ShfCml(L)
    If L = "" Then Exit Function
Wend
End Function

Sub VcMthDotCmlGpAsetInVbe()
MthDotCmlGpAsetInVbe.Srt.Vc
End Sub

Private Sub Z_CmlAy()
Dim Ny$(): Ny = MthNyV
Dim N
For Each N In Ny
    If N <> Jn(CmlAy(CStr(N))) Then Stop
Next
End Sub

Function CmlRel(Ny$()) As Rel
Set CmlRel = Rel(CmlssAy(Ny))
End Function

