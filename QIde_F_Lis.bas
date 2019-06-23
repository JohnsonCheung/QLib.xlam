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
Dim A As Drs: A = DMthPFun
BrwDrs ColPfx(A, "Mthn", Pfx)
End Sub
Sub LisPFunPatn(Patn$)
Dim A As Drs: A = DMthPFun
BrwDrs DwPatn(A, "Mthn", Patn)
End Sub
Sub LisPubPatn(Patn$)
Dim A As Drs: A = DMthPub
BrwDrs DwPatn(A, "Mthn", Patn)
End Sub
Function DMthPub() As Drs
DMthPub = DwEqExl(DMthP, "Mdy", "Pub")
End Function

Function DMthPFun() As Drs
DMthPFun = DwEqExl(DMthPub, "Ty", "Fun")
End Function

Function DMthPatn(Patn$) As Drs
DMthPatn = DwPatn(DMthP, "Mthn", Patn)
End Function

Function DMth2Patn(Patn1$, Patn2$) As Drs
DMth2Patn = Dw2Patn(DMthP, "Mthn", Patn1, Patn2)
End Function

Function DMthoRetAs() As Drs
Dim D As Drs: D = DMthP
Dim I%: I = IxzAy(D.Fny, "MthLin")
Dim Dr, Dy(): For Each Dr In D.Dy
    Dim MthLin$: MthLin = Dr(I)
    Dim RetAs$: RetAs = RetAszL(MthLin)
    PushI Dr, RetAs
    PushI Dy, Dr
Next
DMthoRetAs = DrszAddFF(D, "RetAs", Dy)
End Function

Function DMthRetAsPatn(RetAsPatn$) As Drs
Dim D As Drs: D = DMthoRetAs
DMthRetAsPatn = DwPatn(D, "RetAs", RetAsPatn)
End Function

Function DMthPrp() As Drs
DMthPrp = DwIn(DMthPub, "Ty", SyzSS("Get Let Set"))
End Function
Sub LisPFunRetAs(RetAsPatn$)
Dim RetAs As Drs: RetAs = AddColzRetAs(DMthPFun)
Dim Patn As Drs: Patn = DwPatn(RetAs, "RetAs", RetAsPatn)
Dim T50 As Drs: T50 = TopN(Patn)
BrwDrs T50
End Sub
Sub LisRetAs(RetAsPatn$, Optional N = 50)
Dim RetAs As Drs: RetAs = AddColzRetAs(DMthP)
Dim Patn As Drs: Patn = DwPatn(RetAs, "RetAs", RetAsPatn)
Dim T50 As Drs: T50 = TopN(Patn, N:=N)
BrwDrs T50
End Sub

Sub LisPPrpRetAs(RetAsPatn$)
Dim S As Drs: S = DMthP
Dim RetAs As Drs: RetAs = AddColzRetAs(S)
Dim Pub As Drs: Pub = DwEqExl(RetAs, "Mdy", "Pub")
Dim Fun As Drs: Fun = DwEqExl(Pub, "Ty", "Get")
Dim Patn As Drs: Patn = DwPatn(Fun, "RetAs", RetAsPatn)
Dim T50 As Drs: T50 = TopN(Patn)
BrwDrs T50
End Sub

Sub ListMthRetAs(Patn$)
End Sub

Sub LisMth()
Dim Ay$(): Stop: ' Ay = QMthNyzV(CVbe)
Debug.Print "Fst 30 of " & Si(Ay) & " methods"
D FstNEle(Ay, 30)
End Sub

Sub LisMthPfx(Pfx$, Optional PubOnly As Boolean)
'D QMthNyV(WhStrzPfx(Pfx, PubOnly))
End Sub

Sub LisMthSfx(Sfx$, Optional PubOnly As Boolean)
'D QMthNyV(WhMthzSfx(Sfx, PubOnly))
End Sub

Sub LisMthPatn(Patn$)
DmpDrs DTopN(DMthPatn(Patn))
End Sub

Sub LisMthPatnDrs(Patn$)
Dmp FmtDrszRdu(DwPatn(DMthRetAsPatn("Drs"), "Mthn", Patn))
End Sub

Sub LisMthRetAs(RetAsPatn$, Optional MthPatn$ = ".+")
DmpDrs DTopN(Dw2Patn(DMthoRetAs, "RetAs Mthn", RetAsPatn, MthPatn))
End Sub

Sub BrwMthRetAs(RetAsPatn$)
BrwDrs DwPatn(DMthoRetAs, "RetAs", RetAsPatn)
End Sub

Sub LisMth2Patn(Patn1$, Patn2$)
DmpDrs DTopN(DMth2Patn(Patn1, Patn2))
End Sub

