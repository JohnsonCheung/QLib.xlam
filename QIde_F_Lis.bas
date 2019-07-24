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

Sub LisMth(Optional Patn$, Optional Patn2$, Optional Pub As Boolean, Optional MdnPatn$, Optional RetAsPatn$, _
Optional NPm% = -1, Optional FstPmDclSfx$, Optional FstPmNmPatn$, _
Optional Oup As EmOupTy, Optional NTop% = 50)
Dim IPub  As Drs:        If Pub Then IPub = DoMthP Else IPub = DoPubMth
Dim IPatn As Drs:        If Patn = "" Then IPatn = IPub Else IPatn = DwPatn(IPub, "Mthn", Patn)
Dim IPatn2 As Drs:       If Patn2 = "" Then IPatn2 = IPatn Else IPatn = DwPatn(IPatn2, "Mthn", Patn2)
Dim IPm As Drs:    IPm = AddColzMthPm(IPatn2)
Dim ILin As Drs:  ILin = SelDrsAtEnd(IPm, "MthLin")
Dim IMdn As Drs:         If MdnPatn = "" Then IMdn = ILin Else IMdn = DwPatn(ILin, "Mdn", MdnPatn)
Dim IRet As Drs:         IRet = LisMth__RetAs(IMdn, RetAsPatn)
Dim INPm As Drs:         INPm = LisMth__NPm(IRet, NPm)
Dim IFstPmNm As Drs:     IFstPmNm = LisMth__FstPmNm(IRet, FstPmNmPatn)
Dim IFstPmDclSfx As Drs: IFstPmDclSfx = LisMth__FstPmDclSfx(IFstPmNm, FstPmDclSfx)
Dim ITop As Drs:  ITop = TopN(IFstPmDclSfx, NTop)
Dmp FmtDrszRdu(ITop, , , , EiBeg1, EiSSFmt), Oup
End Sub

Private Function LisMth__FstPmNm(D As Drs, FstPmNmPatn$) As Drs
If FstPmNmPatn = "" Then LisMth__FstPmNm = D: Exit Function
Dim Ix%: Ix = IxzAy(D.Fny, "MthPm")
Dim Re As RegExp: Set Re = RegExp(FstPmNmPatn)
Dim FstPmNm$, ODy(), Dr, Pm$: For Each Dr In Itr(D.Dy)
    Pm = Dr(Ix)
    FstPmNm = BefOrAll(Pm, ",")
    If Re.Test(FstPmNm) Then PushI ODy, Dr
Next
LisMth__FstPmNm = Drs(D.Fny, ODy)
End Function
Private Function LisMth__FstPmDclSfx(D As Drs, Sfx$) As Drs
If Sfx = "" Then LisMth__FstPmDclSfx = D: Exit Function
Dim Ix%: Ix = IxzAy(D.Fny, "Pm")
Dim Re As RegExp: Set Re = RegExp(Sfx)
Dim FstPmNm$, ODy(), Dr, Pm$: For Each Dr In Itr(D.Dy)
    Pm = Dr(Ix)
    FstPmNm = BefOrAll(Pm, ",")
    If DclSfx(FstPmNm) = Sfx Then
        PushI ODy, Dr
    End If
Next
LisMth__FstPmDclSfx = Drs(D.Fny, ODy)
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
Dim I%: I = IxzAy(D.Fny, "MthLin")
Dim Dr, Dy(): For Each Dr In D.Dy
    Dim MthLin$: MthLin = Dr(I)
    Dim RetAs$: RetAs = RetAszL(MthLin)
    PushI Dr, RetAs
    PushI Dy, Dr
Next
LisMth__RetAs = AddColzFFDy(D, "RetAs", Dy)
End Function

