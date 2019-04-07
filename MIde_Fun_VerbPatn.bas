Attribute VB_Name = "MIde_Fun_VerbPatn"
Option Explicit
Public Const C_VerbSs$ = "Zip Wrt Wrp Wait Vis Vc UnderLin UnRmk UnEsc Trim Tile Thw Tak Syw Sye Swap Sum Stop Srt Split Solve Shw Shf Set Sel Sav Run Rpl Rmv Rmk Rfh Reverse Resz Ren ReSz ReSeq ReOrd RTrim Quote Quit Push Prompt Pop Opn Nxt Norm New Mov Mk Minus Min Mid Mge Max Map Lnk Lis Lik Las Kill Jn Jmp Is Into Intersect Ins Initialize Init Inf Indent Inc Imp Hit Has Halt Gen Fst Fmt Flat Fill Extend Expand Exp Exl Evl Esc Ens EndTrim Edt Dryw Drye Drsw Drse Drp Down Do Dmp Dlt Cv Cut Crt Cpy Compress Cls Clr Clone Cln Clear Chk3 Chk2 Chk1 Chk Chg Change Cfm Brw Brk Box Bld Bet Below Bef Bdr Backup Ayw Aye AutoFit AutoExec Ass Asg And Align Aft Add Above"
Public Const C_BRKCmlss$ = "Wi Wo By Of To"
Public Const DocOfQBNm$ = "Quote-Brk-Nm.  If the Cml is BRKCml, quote-bkt."
Property Get BRKCmlASet() As Aset
Static X As Aset
If IsNothing(X) Then Set X = AsetzSsl(C_BRKCmlss)
Set BRKCmlASet = X
End Property
Property Get MthVNyOfVbe() As String()
Dim MthNm
For Each MthNm In Itr(MthNyOfVbe)
    PushI MthVNyOfVbe, MthVNm(MthNm)
Next
End Property
Private Sub Z_MthVNsetOfVbe()
MthVNsetOfVbe.Srt.Vc
End Sub
Property Get MthDNmToMdDNmRelOfVbe() As Rel
Set MthDNmToMdDNmRelOfVbe = MthDNmToMdDNmRelzVbe(CurVbe)
End Property

Private Function MthDNmToMdDNmRelzVbe(A As Vbe) As Rel
Set MthDNmToMdDNmRelzVbe = RelzDotLy(MthQNyzVbe(A))
End Function
Property Get MthVNsetOfVbe() As Aset
Set MthVNsetOfVbe = AsetzAy(MthVNyOfVbe)
End Property

Function MthQVNsetOfVbe(Optional WhStr$) As Aset
Dim Ay$(): Ay = MthQVNyOfVbe(WhStr)
Set MthQVNsetOfVbe = AsetzAy(Ay)
End Function

Sub VcMthQVNsetOfVbe(Optional WhStr$)
MthQVNsetOfVbe(WhStr).Srt.Vc
End Sub

Sub VcMthQVNyOfVbe(Optional WhStr$)
Vc AyQSrt(MthQVNyOfVbe(WhStr))
End Sub

Function MthQVNyOfVbe(Optional WhStr$) As String()
MthQVNyOfVbe = MthQVNyzVbe(CurVbe, WhStr)
End Function

Function MthQVNyzVbe(A As Vbe, Optional WhStr$) As String()
MthQVNyzVbe = QVNy(MthNyzVbe(A, WhStr))
End Function

Function QVNy(Ny$()) As String()
Dim Nm
For Each Nm In Itr(Ny)
    PushI QVNy, QVNm(Nm)
Next
End Function

Function QBNm$(Nm)
Dim Cml, O$()
For Each Cml In Itr(Cml1Ay(Nm))
    If IsBRKCml(Cml) Then
        PushI O, QuoteBkt(Cml)
    Else
        PushI O, Cml
    End If
Next
QBNm = Jn(O)
End Function

Function QVBNm$(Nm) 'Quote-Verb-and-cmlBrk-Nm.
Dim V$: V = Verb(Nm)
If V = "" Then
    QVBNm = "#" & QBNm(Nm)
Else
    With Brk(Nm, V)
    QVBNm = QBNm(.S1) & QuoteSq(V) & QBNm(.S2)
    End With
End If
End Function
Function QVNm$(Nm)
Dim V$: V = Verb(Nm)
If V = "" Then
    QVNm = "#" & Nm
Else
    QVNm = Replace(Nm, V, QuoteBkt(V), Count:=1)
End If
End Function
Function MthVNm$(MthNm)
MthVNm = Verb(MthNm) & "." & MthNm
End Function
Property Get VerbRe() As RegExp
Static X As RegExp
If IsNothing(X) Then Set X = RegExp(PatnzVerbss(C_VerbSs))
Set VerbRe = X
End Property
Sub BrwVerb()
Vc SySsl(C_VerbSs)
End Sub
Sub VcNVTDNmAsetOfVbe()
NVTDNmAsetOfVbe.Srt.Vc
End Sub
Property Get NVTDNmAsetOfVbe() As Aset
Set NVTDNmAsetOfVbe = AsetzAy(NVTDNyOfVbe)
End Property
Property Get NVTDNyOfVbe() As String()
NVTDNyOfVbe = NVTDNyzVbe(CurVbe)
End Property
Private Function NVTDNyzVbe(A As Vbe) As String()
NVTDNyzVbe = NVTDNy(MthNyzVbe(A))
End Function
Private Function NVTDNy(Ny$()) As String()
Dim Nm
For Each Nm In Itr(Ny)
    PushI NVTDNy, NVTDNm(Nm)
Next
End Function
Private Function NVTDNm$(Nm) 'Nm.Verb.Ty.Dot-Nm
NVTDNm = NVTy(Nm) & "." & Nm
End Function
Function FstVerbSubNyOfVbe() As String()

End Function
Function NVTy$(Nm) 'Nm.Verb-Ty
Select Case True
Case IsNoVerbNm(Nm): NVTy = "NoVerb"
Case IsFstVerbNm(Nm): NVTy = "FstVerb"
Case IsMidVerbNm(Nm): NVTy = "MidVerb"
Case Else: Thw CSub, "Program error: a Nm must be any of [NoVerb | FstVerb | MidVerb]", "Nm", Nm
End Select
End Function
Function IsNoVerbNm(Nm) As Boolean
IsNoVerbNm = Verb(Nm) = ""
End Function
Function IsMidVerbNm(Nm) As Boolean
Dim V$: V = Verb(Nm): If V = "" Then Exit Function
IsMidVerbNm = Not HasPfx(Nm, Verb(Nm))
End Function

Function IsFstVerbNm(Nm) As Boolean
IsFstVerbNm = HasPfx(Nm, Verb(Nm))
End Function
Function IsVerb(S) As Boolean
IsVerb = VerbAset.Has(RmvEndDig(S))
End Function

Property Get VerbAset() As Aset
Static X As Aset
If IsNothing(X) Then Set X = AsetzSsl(C_VerbSs)
Set VerbAset = X
End Property
Function RmvEndDig$(S)
Dim J&
For J = Len(S) To 1 Step -1
    If Not IsAscDig(Asc(Mid(S, J, 1))) Then
        RmvEndDig = Left(S, J)
        Exit Function
    End If
Next
RmvEndDig = Left(S, J)
End Function
Function Verb$(Nm)
Dim Cml, LetterCml$
For Each Cml In Cml1Ay(Nm)
    LetterCml = RmvDigSfx(Cml)
    If VerbAset.Has(LetterCml) Then Verb = Cml: Exit Function
Next
End Function
Property Get NormVerbss$()
NormVerbss = NormSsl(C_VerbSs, IsDes:=True)
End Property
Function NormSsl$(Ssl, Optional IsDes As Boolean)
NormSsl = JnSpc(AyQSrt(AywDist(SySsl(Ssl)), IsDes:=True))
End Function

Function PatnzVerbss$(Verbss$)
Dim O$(), Verb
For Each Verb In AsetzAy(SySsl(Verbss)).Itms
    PushI O, PatnzVerb(Verb)
Next
PatnzVerbss = QuoteBkt(JnVbar(O))
End Function

Private Function PatnzVerb$(Verb)
ThwIfNotVerb Verb, CSub
PatnzVerb = Verb & "[^a-z|0-9]*"
End Function
Private Sub ThwIfNotVerb(S, Fun$)
If Not IsNm(S) Then Thw Fun, "Verb must be a name", "Str", S
If Not IsAscUCas(Asc(FstChr(S))) Then Thw Fun, "Verb must started with UCase", "Str", S
End Sub

Function QuoteVerb$(Nm)

End Function
