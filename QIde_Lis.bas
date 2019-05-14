Attribute VB_Name = "QIde_Lis"
Option Explicit
Private Const CMod$ = "MIde_Lis."
Private Const Asm$ = "QIde"
Function MdzPjLisDt(P As VBProject, Optional B As WhMd) As Dt
Stop '
End Function

Sub MdzPjLisBrwDt(P As VBProject, Optional B As WhMd)
BrwDt MdzPjLisDt(P, B)
End Sub

Sub MdzPjLisDmpDt(P As VBProject, Optional B As WhMd)
DmpDt MdzPjLisDt(P, B)
End Sub

Sub LisMd(Optional Patn$, Optional Exl$)
Dim A$()
'    A = CmpNyPj(CPj, WhMd("Std", WhNm(Patn, Exl)))
    A = AySrt(A)
    A = AddPfxzAy(A, "ShwMbr """)
D A
End Sub
Sub LisPj()
Dim A$()
    A = PjNyzV(CVbe)
    D AddPfxzAy(A, "ShwPj """)
D A
End Sub

Sub LisStopLin()

End Sub
Sub LisMth(Optional WhStr$)
Dim Ay$(): Ay = MthQNyzV(CVbe, WhStr)
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
D MthQNyV(WhStrzPfx(Pfx, PubOnly))
End Sub

Sub LisMthSfx(Sfx$, Optional PubOnly As Boolean)
D MthQNyV(WhMthzSfx(Sfx, PubOnly))
End Sub

Sub LisMthPatn(Patn$, Optional InclPrv As Boolean)
D MthQLyV(WhStrzMthPatn(Patn, InclPrv))
End Sub
