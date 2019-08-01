Attribute VB_Name = "QIde_F_Lis"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Lis."
Private Const Asm$ = "QIde"

Sub LisCmpzMd(Optional Patn$, Optional Exl$)
Dim A$()
'    A = CmpNyPj(CPj, WhMd("Std", WhNm(Patn, Exl)))
    A = SrtAy(A)
    A = AddPfxzAy(A, "ShwMbr """)
D A
End Sub
Sub LisPj()
Dim A$()
    A = PjnyzV(CVbe)
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
DoPubSub = DwEq(DoPubMth, "Ty", "Sub")
End Function

Sub LisPubPatn(Patn$)
Dim A As Drs: A = DoPubMth
BrwDrs DwPatn(A, "Mthn", Patn)
End Sub

Function DoPubFun() As Drs
DoPubFun = DwEqExl(DoPubMth, "Ty", "Fun")
End Function

Function DoPubMthPatn(Patn$) As Drs
DoPubMthPatn = DwPatn(DoPubMth, "Mthn", Patn)
End Function

Function DoMthRetAsPatn(RetAsPatn$) As Drs
Dim D As Drs: D = DoMthRetAs
DoMthRetAsPatn = DwPatn(D, "RetAs", RetAsPatn)
End Function

Function DoPubPrp() As Drs
DoPubPrp = DwIn(DoPubMth, "Ty", SyzSS("Get Let Set"))
End Function

Sub LisPFunRetAs(RetAsPatn$)
Dim RetAs As Drs: RetAs = AddColzRetAs(DoPubMth)
Dim Patn As Drs: Patn = DwPatn(RetAs, "RetAs", RetAsPatn)
Dim T50 As Drs: T50 = TopN(Patn)
BrwDrs T50
End Sub
Sub LisRetAs(RetAsPatn$, Optional N = 50)
Dim RetAs As Drs: RetAs = AddColzRetAs(DoPubMth)
Dim Patn As Drs: Patn = DwPatn(RetAs, "RetAs", RetAsPatn)
Dim T50 As Drs: T50 = TopN(Patn, N:=N)
BrwDrs T50
End Sub

Sub LisPPrpRetAs(RetAsPatn$)
Dim S As Drs: S = DoPubMth
Dim RetAs As Drs: RetAs = AddColzRetAs(S)
Dim Pub As Drs: Pub = DwEqExl(RetAs, "Mdy", "Pub")
Dim Fun As Drs: Fun = DwEqExl(Pub, "Ty", "Get")
Dim Patn As Drs: Patn = DwPatn(Fun, "RetAs", RetAsPatn)
Dim T50 As Drs: T50 = TopN(Patn)
BrwDrs T50
End Sub

Sub LisMthCntzQIde()
DmpDrs SrtDrs(DwEq(DoMthCntP, "Lib", "QIde")), Fmt:=EiSSFmt, IsSum:=True
End Sub

Private Function JSrc$(Mdn$, Lno&, Lin)
':JSrc: :Lin #Jmp-Lin# ! Fmt: T1 Rst, *T1 is JmpLin"<mdn:Lno>".  *Rst is '<SrcLin>
JSrc = FmtQQ("JmpLin""?:?"" '?", Mdn, Lno, Lin)
End Function
Function JSrczPred(P As IPred) As String()
Dim O$()
Dim C As VBComponent: For Each C In CPj.VBComponents
    Dim Md As CodeModule: Set Md = C.CodeModule
    Dim L, Lno&: Lno = 0
    If Md.CountOfLines > 0 Then
        For Each L In Itr(SplitCrLf(Md.Lines(1, Md.CountOfLines)))
            Lno = Lno + 1
            If P.Pred(L) Then
                PushI O, JSrc(C.Name, Lno, L)
            End If
        Next
    End If
Next
JSrczPred = AlignLyzTRst(O)
End Function

Private Function JSrczPfx(LinPfx$) As String()
JSrczPfx = JSrczPred(PredHasPfx(LinPfx))
End Function

Private Function JSrczPatn(LinPatn$) As String()
':JSrc: :CdLy #Jmp-Src# ! It is
JSrczPatn = JSrczPred(PredHasPatn(LinPatn))
End Function

Sub LisSrcPfx(LinPfx$)
Dmp JSrczPfx(LinPfx)
End Sub

Sub LisSrc(LinPatn$)
Dmp JSrczPatn(LinPatn)
End Sub


Sub LisMth(Optional Patn$, Optional Patn2$, Optional Pub As Boolean, Optional MdnPatn$, Optional RetAsPatn$, _
Optional NPm% = -1, Optional ArgPatn$, _
Optional Oup As EmOupTy, Optional NTop% = 50)
Dim D As Drs: D = DoMthLis(Patn, Patn2, Pub, MdnPatn, RetAsPatn, NPm, ArgPatn)
Dim D1 As Drs: D1 = TopN(D, NTop)
Dmp FmtDrszRdu(D1, , , , EiBeg1, EiSSFmt), Oup
End Sub

Function DoMthLis(Optional Patn$, Optional Patn2$, Optional Pub As Boolean, Optional MdnPatn$, Optional RetAsPatn$, _
Optional NPm% = -1, Optional ArgPatn$) As Drs
Dim IPub  As Drs:        If Pub Then IPub = DoPubMth Else IPub = DoMthP
Dim IPatn As Drs:        If Patn = "" Then IPatn = IPub Else IPatn = DwPatn(IPub, "Mthn", Patn)
Dim IPatn2 As Drs:       If Patn2 = "" Then IPatn2 = IPatn Else IPatn2 = DwPatn(IPatn, "Mthn", Patn2)
Dim IPm As Drs:    IPm = AddColzMthPm(IPatn2)
Dim ILin As Drs:  ILin = SelDrsAtEnd(IPm, "MthLin")
Dim IMdn As Drs:         If MdnPatn = "" Then IMdn = ILin Else IMdn = DwPatn(ILin, "Mdn", MdnPatn)
Dim IRet As Drs:         IRet = LisMth__RetAs(IMdn, RetAsPatn)
Dim INPm As Drs:         INPm = LisMth__NPm(IRet, NPm)
Dim IArgPatn As Drs:     IArgPatn = LisMth__ArgPatn(INPm, ArgPatn)
DoMthLis = IArgPatn
End Function

Private Function LisMth__ArgPatn(D As Drs, ArgPatn$) As Drs
If ArgPatn = "" Then LisMth__ArgPatn = D: Exit Function
Dim Ix%: Ix = IxzAy(D.Fny, "MthPm", ThwEr:=EiThwEr)
Dim Re As RegExp: Set Re = RegExp(ArgPatn)
Dim ArgAy$(), ODy(), Dr: For Each Dr In Itr(D.Dy)
    ArgAy = SplitCommaSpc(Dr(Ix))
    If HasEleRe(ArgAy, Re) Then PushI ODy, Dr
Next
LisMth__ArgPatn = Drs(D.Fny, ODy)
End Function

Private Function LisMth__NPm(D As Drs, NPm%) As Drs
If NPm < 0 Then LisMth__NPm = D: Exit Function
Dim Ix%: Ix = IxzAy(D.Fny, "Pm")
Dim ODy(), Dr, Pm$: For Each Dr In Itr(D.Dy)
    Pm = Dr(Ix)
    If Si(SplitComma(Pm)) = NPm Then PushI ODy, Dr
Next
LisMth__NPm = Drs(D.Fny, ODy)
End Function

Private Function LisMth__RetAs(D As Drs, RetAsPatn$) As Drs
If RetAsPatn = "" Then LisMth__RetAs = D: Exit Function
Dim R As RegExp: Set R = RegExp(RetAsPatn)
Dim I%: I = IxzAy(D.Fny, "MthLin")
Dim Dr, Dy(): For Each Dr In D.Dy
    Dim MthLin$: MthLin = Dr(I)
    Dim RetAs$: RetAs = RetAszL(MthLin)
    If R.Test(RetAs) Then
        PushI Dr, RetAs
        PushI Dy, Dr
    End If
Next
LisMth__RetAs = AddColzFFDy(D, "RetAs", Dy)
End Function

