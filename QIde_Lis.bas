Attribute VB_Name = "QIde_Lis"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Lis."
Private Const Asm$ = "QIde"

Sub LIsCmpzMd(Optional Patn$, Optional Exl$)
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
Sub LisPFunRetAs(RetAsPatn$)
Dim S As Drs: S = DMthP
Dim RetAs As Drs: RetAs = AddColzRetAs(S)
Dim Pub As Drs: Pub = ColEqE(RetAs, "Mdy", "Pub")
Dim Fun As Drs: Fun = ColEqE(Pub, "Ty", "Fun")
Dim Patn As Drs: Patn = ColPatn(Fun, "RetAs", RetAsPatn)
Dim T50 As Drs: T50 = TopN(Patn)
BrwDrs T50
End Sub
Sub LisRetAs(RetAsPatn$, Optional N = 50)
Dim S As Drs: S = DMthP
Dim RetAs As Drs: RetAs = AddColzRetAs(S)
Dim Patn As Drs: Patn = ColPatn(RetAs, "RetAs", RetAsPatn)
Dim T50 As Drs: T50 = TopN(Patn, N:=N)
BrwDrs T50
End Sub

Sub LisPPrpRetAs(RetAsPatn$)
Dim S As Drs: S = DMthP
Dim RetAs As Drs: RetAs = AddColzRetAs(S)
Dim Pub As Drs: Pub = ColEqE(RetAs, "Mdy", "Pub")
Dim Fun As Drs: Fun = ColEqE(Pub, "Ty", "Get")
Dim Patn As Drs: Patn = ColPatn(Fun, "RetAs", RetAsPatn)
Dim T50 As Drs: T50 = TopN(Patn)
BrwDrs T50
End Sub
Sub LisMth()
Dim Ay$(): Stop: ' Ay = QMthnyzV(CVbe)
Debug.Print "Fst 30 of " & Si(Ay) & " methods"
D AywFstNEle(Ay, 30)
End Sub

Private Function WhStrzMthPatn$(MthPatn$, Optional PubOnly As Boolean)
WhStrzMthPatn = " -MthPatn " & MthPatn & WhStrzPubOnly(PubOnly)
End Function

Private Function WhStrzPubOnly$(PubOnly As Boolean)
If PubOnly Then WhStrzPubOnly = " -Pub"
End Function

Function WhStrzPfx$(MthPfx$, Optional PubOnly As Boolean)
WhStrzPfx = WhStrzMthPatn("^" & MthPfx, PubOnly)
End Function

Function WhStrzSfx$(MthSfx$, Optional PubOnly As Boolean)
WhStrzSfx = WhStrzMthPatn(MthSfx & "$", PubOnly)
End Function

Sub LisMthPfx(Pfx$, Optional PubOnly As Boolean)
'D QMthnyV(WhStrzPfx(Pfx, PubOnly))
End Sub

Sub LisMthSfx(Sfx$, Optional PubOnly As Boolean)
'D QMthnyV(WhMthzSfx(Sfx, PubOnly))
End Sub

Sub LisMthPatn(Patn$, Optional InclPrv As Boolean)
'D MthQLyV(WhStrzMthPatn(Patn, InclPrv))
End Sub
