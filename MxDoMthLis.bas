Attribute VB_Name = "MxDoMthLis"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDoMthLis."
Enum EmTriSte
    EiTriOpn
    EiTriYes
    EiTriNo
End Enum
Public Const FFoMthLis$ = "Pjn MdTy Mdn L Mdy Ty Mthn TyChr RetAs ShtPm"
':JSrc: :Lin #Jmp-Lin# ! Fmt: T1 Rst, *T1 is JmpLin"<mdn:Lno>".  *Rst is '<SrcLin>

Sub LisPj()
Dim A$()
    A = PjNyzV(CVbe)
    D AddPfxzAy(A, "ShwPj """)
D A
End Sub

Sub LisStopLin()

End Sub
Sub LisPFunPfx(Pfx$)
Dim A As Drs: A = DoPubPrp
BrwDrs ColPfx(A, "Mthn", Pfx)
End Sub

Sub LisPubSub(MthPatn$)
DmpDrs DwPatn(DoPubSub, "Mthn", MthPatn)
End Sub

Sub LisPubZ()
Dim A As Drs: A = DwPatn(DoPubSub, "Mthn", "^Z$")
Dmp AddSfxzAy(StrCol(A, "Mdn"), ".Z")
End Sub

Sub LisPubFun(MthPatn$)
DmpDrs DwPatn(DoPubFun, "Mthn", MthPatn)
End Sub

Function DoPubSub() As Drs
DoPubSub = DwEq(DoPubFun, "Ty", "Sub")
End Function

Sub LisPubPatn(Patn$)
Dim A As Drs: A = DoPubFun
BrwDrs DwPatn(A, "Mthn", Patn)
End Sub

Function DoPubFun() As Drs
DoPubFun = DwEqExl(DoPubFun, "Ty", "Fun")
End Function

Function DoPubFunPatn(Patn$) As Drs
DoPubFunPatn = DwPatn(DoPubFun, "Mthn", Patn)
End Function

Function DoPubPrp() As Drs
DoPubPrp = DwIn(DoPubFun, "Ty", SyzSS("Get Let Set"))
End Function

Sub LisPFunRetAs(RetAsPatn$)
Dim RetSfx As Drs: RetSfx = AddColzRetAs(DoPubFun)
Dim Patn As Drs: Patn = DwPatn(RetSfx, "RetSfx", RetAsPatn)
Dim T50 As Drs: T50 = DwTopN(Patn)
BrwDrs T50
End Sub

Sub LisRetAs(RetAsPatn$, Optional N = 50)
Dim RetSfx As Drs: RetSfx = AddColzRetAs(DoPubFun)
Dim Patn As Drs: Patn = DwPatn(RetSfx, "RetSfx", RetAsPatn)
Dim T50 As Drs: T50 = DwTopN(Patn, N:=N)
BrwDrs T50
End Sub

Sub LisPPrpRetAs(RetAsPatn$)
Dim S As Drs: S = DoPubFun
Dim RetSfx As Drs: RetSfx = AddColzRetAs(S)
Dim Pub As Drs: Pub = DwEqExl(RetSfx, "Mdy", "Pub")
Dim Fun As Drs: Fun = DwEqExl(Pub, "Ty", "Get")
Dim Patn As Drs: Patn = DwPatn(Fun, "RetSfx", RetAsPatn)
Dim T50 As Drs: T50 = DwTopN(Patn)
BrwDrs T50
End Sub

Sub LisMthCntzQIde()
DmpDrs SrtDrs(DwEq(DoMthCntP, "Lib", "QIde")), Fmt:=EiSSFmt, IsSum:=True
End Sub

Private Function JSrc$(Mdn$, Lno&, C1%, C2%, Lin)
JSrc = FmtQQ("JmpLin""?:?:?:?"" '?", Mdn, Lno, C1, C2, Lin)
End Function

Function JSrczPred(P As IPred) As String()
Dim O$()
Dim C As VBComponent: For Each C In CPj.VBComponents
    Dim Md As CodeModule: Set Md = C.CodeModule
    Dim L, Lno&: Lno = 0
    Dim C1%, C2%
    If Md.CountOfLines > 0 Then
        For Each L In Itr(SplitCrLf(Md.Lines(1, Md.CountOfLines)))
            Lno = Lno + 1
            If P.Pred(L) Then
                C1 = 1
                C2 = 1
                PushI O, JSrc(C.Name, Lno, C1, C2, L)
            End If
        Next
    End If
Next
JSrczPred = AlignLyzTRst(O)
End Function

Private Function JSrczIdf(Idf$) As String()
JSrczIdf = JSrczPred(PredHasIdf(Idf))
End Function

Private Function JSrczPfx(LinPfx$) As String()
JSrczPfx = JSrczPred(PredHasPfx(LinPfx))
End Function

Private Function JSrczPatn(LinPatn$, Optional AndPatn1$, Optional AndPatn2$) As String()
JSrczPatn = JSrczPred(PredHasPatn(LinPatn, AndPatn1, AndPatn2))
End Function

Sub LisSrcoPfx(LinPfx$, Optional OupTy As EmOupTy = EmOupTy.EiOtDmp)
Brw JSrczPfx(LinPfx), OupTy:=OupTy
End Sub

Sub LisSrcoIdf(Idf$, Optional OupTy As EmOupTy = EmOupTy.EiOtDmp)
':Idf: :Nm #Identifier#
Brw JSrczIdf(Idf), OupTy:=OupTy
End Sub

Sub LisSrc(LinPatn$, Optional AndPatn1$, Optional AndPatn2$, Optional OupTy As EmOupTy = EmOupTy.EiOtDmp)
Brw JSrczPatn(LinPatn, AndPatn1, AndPatn2), OupTy:=OupTy
End Sub

Sub LisMth(Optional Patn$, Optional Patn1$, Optional Patn2$, Optional ShtMdySS$, Optional ShtMthTySS$, Optional MdnPatn$, Optional TyChr$, Optional RetAsPatn$, _
Optional NPm% = -1, Optional ShtPmPatn$, Optional HasAp As EmTriSte, Optional RetAy As EmTriSte, _
Optional OupTy As EmOupTy, Optional Top% = 50)
Dim D As Drs: D = SelDoMthLis(DoMthLisP, Patn, Patn1, Patn2, ShtMdySS, ShtMthTySS, MdnPatn, TyChr, RetAsPatn, NPm, ShtPmPatn, HasAp, RetAy)
Dim D1 As Drs: D1 = DwTopN(D, Top)
Brw FmtCellDrszRdu(D1, , , , EiBeg1, EiSSFmt), OupTy:=OupTy
End Sub

Function PatnzSS$(SS, LisAy$())
Dim A$(): A = AwDist(SyzSS(SS))
Dim B$()
    Dim I: For Each I In Itr(A)
        If HasEle(LisAy, I) Then
            PushNDup B, I
        End If
    Next
Dim C$: C = Jn(B, "|")
If C = "" Then Exit Function
PatnzSS = Qte(C, "()")
End Function

Function ßAA()
End Function

Function DoMthLisP() As Drs
Static X As Boolean, Y As Drs
If X Then GoTo XX
X = True
Dim ITyChr   As Drs:   ITyChr = AddMthColTyChr(DoMthP)
Dim IPm      As Drs:      IPm = AddMthColMthPm(ITyChr)
Dim IShtPm   As Drs:   IShtPm = AddMthColShtPm(IPm)
Dim IRetAs   As Drs:   IRetAs = AddColoRetAs(IShtPm)
                            Y = SelDrs(IRetAs, FFoMthLis)
XX:                 DoMthLisP = Y
End Function

Private Function SelDoMthLis(DoMthLis As Drs, Patn$, Patn1$, Patn2$, ShtMdySS$, ShtMthTySS$, MdnPatn$, TyChr$, RetAsPatn$, _
NPm%, ShtPmPatn$, HasAp As EmTriSte, RetAy As EmTriSte) As Drs
'- Pfx-Pn = Patn
Dim PnMdy$:             PnMdy = PatnzSS(ShtMdySS, ShtMthMdyAy)
Dim PnTy$:               PnTy = PatnzSS(ShtMthTySS, ShtMthTyAy)
'- Pfx-I = Inp-Do-Fm-DoMthLis
Dim IMdy     As Drs:     IMdy = DwPatn(DoMthLisP, "Mdy", PnMdy)
Dim ITy      As Drs:      ITy = DwPatn(IMdy, "Ty", PnTy)
Dim ITyChr As Drs:     ITyChr = DwEqStr(ITy, "TyChr", TyChr)
Dim IPatn    As Drs:    IPatn = DwPatn(ITyChr, "Mthn", Patn, Patn1, Patn2)
Dim IHasAp   As Drs:   IHasAp = DwHasAp(IPatn, HasAp)
Dim INPm     As Drs:     INPm = DwNPm(IHasAp, NPm)
Dim IMdn     As Drs:     IMdn = DwPatn(INPm, "Mdn", MdnPatn)
Dim IRetAs   As Drs:   IRetAs = DwPatn(IMdn, "RetAs", RetAsPatn)
Dim IRetAy   As Drs:   IRetAy = DwRetAy(IRetAs, RetAy)
                  SelDoMthLis = DwPatn(IRetAy, "ShtPm", ShtPmPatn)
End Function

Private Function DwRetAy(WiRetAs As Drs, RetAy As EmTriSte) As Drs
If RetAy = EiTriOpn Then DwRetAy = WiRetAs: Exit Function
Dim RetAy1 As Boolean: RetAy1 = BoolzTriSte(RetAy)
Dim IRetAs%: IRetAs = IxzAy(WiRetAs.Fny, "RetAs")
Dim ODy()
    Dim Dr: For Each Dr In Itr(WiRetAs.Dy)
        Dim RetAs$: RetAs = Dr(IRetAs)
        If HasSfx(RetAs, "()") = RetAy1 Then PushI ODy, Dr
    Next
DwRetAy = Drs(WiRetAs.Fny, ODy)
End Function

Private Function HasAp(MthPm) As Boolean
Dim A$(): A = SplitCommaSpc(MthPm): If Si(A) = 0 Then Exit Function
HasAp = HasPfx(LasEle(A), "ParamArray ")
End Function

Function BoolzTriSte(A As EmTriSte) As Boolean
Select Case True
Case A = EiTriYes: BoolzTriSte = True
Case A = EiTriNo:  BoolzTriSte = False
Case Else: Stop
End Select
End Function

Private Function DwHasAp(WiMthPm As Drs, HasAp0 As EmTriSte) As Drs
If HasAp0 = EiTriOpn Then DwHasAp = WiMthPm: Exit Function
Dim HasAp1 As Boolean: HasAp1 = BoolzTriSte(HasAp0)
Dim IMthPm%: IMthPm = IxzAy(WiMthPm.Fny, "MthPm")
Dim ODy()
    Dim Dr: For Each Dr In Itr(WiMthPm.Dy)
        Dim MthPm$: MthPm = Dr(IMthPm)
        If HasAp1 = HasAp(MthPm) Then PushI ODy, Dr
    Next
DwHasAp = Drs(WiMthPm.Fny, ODy)
End Function

Private Function DwNPm(D As Drs, NPm%) As Drs
If NPm < 0 Then DwNPm = D: Exit Function
Dim Ix%: Ix = IxzAy(D.Fny, "MthPm")
Dim ODy(), Dr, Pm$: For Each Dr In Itr(D.Dy)
    Pm = Dr(Ix)
    If Si(SplitComma(Pm)) = NPm Then PushI ODy, Dr
Next
DwNPm = Drs(D.Fny, ODy)
End Function

Private Function AddColoRetAs(WiMthLin As Drs) As Drs
Dim I%: I = IxzAy(WiMthLin.Fny, "MthLin")
Dim Dr, Dy(): For Each Dr In Itr(WiMthLin.Dy)
    Dim MthLin$: MthLin = Dr(I)
    Dim Ret$: Ret = RetAs(MthLin)
    PushI Dr, Ret
    PushI Dy, Dr
Next
AddColoRetAs = AddColzFFDy(WiMthLin, "RetAs", Dy)
End Function