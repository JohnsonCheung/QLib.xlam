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
Dim A As Drs: A = DoMthPFun
BrwDrs ColPfx(A, "Mthn", Pfx)
End Sub
Sub LisPFunPatn(Patn$)
Dim A As Drs: A = DoMthPFun
BrwDrs DwPatn(A, "Mthn", Patn)
End Sub
Sub LisPubPatn(Patn$)
Dim A As Drs: A = DoMthPub
BrwDrs DwPatn(A, "Mthn", Patn)
End Sub
Function DoMthPub() As Drs
DoMthPub = DwEqExl(DoMthP, "Mdy", "Pub")
End Function

Function DoMthPFun() As Drs
DoMthPFun = DwEqExl(DoMthPub, "Ty", "Fun")
End Function

Function DoMthPatn(Patn$) As Drs
DoMthPatn = DwPatn(DoMthP, "Mthn", Patn)
End Function

Function DoMth2Patn(Patn1$, Patn2$) As Drs
DoMth2Patn = Dw2Patn(DoMthP, "Mthn", Patn1, Patn2)
End Function

Function DoMthRetAs() As Drs
Dim D As Drs: D = DoMthP
Dim I%: I = IxzAy(D.Fny, "MthLin")
Dim Dr, Dy(): For Each Dr In D.Dy
    Dim MthLin$: MthLin = Dr(I)
    Dim RetAs$: RetAs = RetAszL(MthLin)
    PushI Dr, RetAs
    PushI Dy, Dr
Next
DoMthRetAs = AddColzFFDy(D, "RetAs", Dy)
End Function

Function DoMthRetAsPatn(RetAsPatn$) As Drs
Dim D As Drs: D = DoMthRetAs
DoMthRetAsPatn = DwPatn(D, "RetAs", RetAsPatn)
End Function

Function DoMthPrp() As Drs
DoMthPrp = DwIn(DoMthPub, "Ty", SyzSS("Get Let Set"))
End Function
Sub LisPFunRetAs(RetAsPatn$)
Dim RetAs As Drs: RetAs = AddColzRetAs(DoMthPFun)
Dim Patn As Drs: Patn = DwPatn(RetAs, "RetAs", RetAsPatn)
Dim T50 As Drs: T50 = TopN(Patn)
BrwDrs T50
End Sub
Sub LisRetAs(RetAsPatn$, Optional N = 50)
Dim RetAs As Drs: RetAs = AddColzRetAs(DoMthP)
Dim Patn As Drs: Patn = DwPatn(RetAs, "RetAs", RetAsPatn)
Dim T50 As Drs: T50 = TopN(Patn, N:=N)
BrwDrs T50
End Sub

Sub LisPPrpRetAs(RetAsPatn$)
Dim S As Drs: S = DoMthP
Dim RetAs As Drs: RetAs = AddColzRetAs(S)
Dim Pub As Drs: Pub = DwEqExl(RetAs, "Mdy", "Pub")
Dim Fun As Drs: Fun = DwEqExl(Pub, "Ty", "Get")
Dim Patn As Drs: Patn = DwPatn(Fun, "RetAs", RetAsPatn)
Dim T50 As Drs: T50 = TopN(Patn)
BrwDrs T50
End Sub

Sub ListMthRetAs(Patn$)
End Sub

Sub LisMth(Optional Patn$ = ".+")
LisMthPatn Patn
End Sub

Sub LisMthPfx(Pfx$, Optional PubOnly As Boolean, Optional MdnPatn$)
LisMthPatn "^" & Pfx, PubOnly, MdnPatn
End Sub

Sub LisMthSfx(Sfx$, Optional PubOnly As Boolean, Optional MdnPatn$)
LisMthPatn Sfx & "$", PubOnly, MdnPatn
End Sub

Sub LisMthPatn(Patn$, Optional PubOnly As Boolean, Optional MdnPatn$)
Dim A As Drs: A = DoMthP
Dim B As Drs: B = DwPatn(A, "Mthn", Patn)
Dim C As Drs: C = AddColzMthPm(B)
Dim D As Drs: D = SelDrsAtEnd(C, "MthLin")
Dim E As Drs:     If MdnPatn = "" Then E = D Else E = DwPatn(D, "Mdn", MdnPatn)
Dim F As Drs: F = TopN(E)
DmpDrszRdu F, Fmt:=EiSSFmt
End Sub

Sub LisMthnRetAs(RetAsPatn$, Optional MthPatn$ = ".+")
Dmp StrCol(Dw2Patn(DoMthRetAs, "RetAs Mthn", RetAsPatn, MthPatn), "Mthn")
End Sub

Sub LisMthRetAs(RetAsPatn$, Optional MthPatn$ = ".+")
Dim A As Drs: A = Dw2Patn(DoMthRetAs, "RetAs Mthn", RetAsPatn, MthPatn)
Dim B As Drs: B = DTopN(A)
DmpDrszRdu B, Fmt:=EiSSFmt
End Sub

Sub BrwMthRetAs(RetAsPatn$)
BrwDrs DwPatn(DoMthRetAs, "RetAs", RetAsPatn)
End Sub

Sub LisMth2Patn(Patn1$, Patn2$)
DmpDrs DTopN(DoMth2Patn(Patn1, Patn2))
End Sub

