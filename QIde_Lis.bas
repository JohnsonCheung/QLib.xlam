Attribute VB_Name = "QIde_Lis"
Option Explicit
Private Const CMod$ = "MIde_Lis."
Private Const Asm$ = "QIde"

Sub MdzPjLisBrwDt(P As VBProject)
BrwDt MdzPjLisDt(P, B)
End Sub

Sub MdzPjLisDmpDt(P As VBProject)
DmpDt MdzPjLisDt(P, B)
End Sub

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
Sub LisMth()
Dim Ay$(): Ay = QMthnyzV(CVbe)
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
D QMthnyV(WhStrzPfx(Pfx, PubOnly))
End Sub

Sub LisMthSfx(Sfx$, Optional PubOnly As Boolean)
D QMthnyV(WhMthzSfx(Sfx, PubOnly))
End Sub

Sub LisMthPatn(Patn$, Optional InclPrv As Boolean)
D MthQLyV(WhStrzMthPatn(Patn, InclPrv))
End Sub
