Attribute VB_Name = "QVb_Lin_Vbl"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Lin_Vbl."
Private Const Asm$ = "QVb"
Function DrszTRst(FF$, TRstLy$()) As Drs
DrszTRst = DrszFF(FF, DyoTRst(TRstLy))
End Function
Function DyoTRst(TRstLy$()) As Variant()
Dim L: For Each L In Itr(TRstLy)
    PushI DyoTRst, SyzTRst(L)
Next
End Function
Function DyoTLiny(TLiny$()) As Variant()
Dim I
For Each I In Itr(TLiny)
    PushI DyoTLiny, TermAy(I)
Next
End Function

Function DyoVblLy(A$()) As Variant()
Dim I
For Each I In Itr(A)
    PushI DyoVblLy, AyTrim(SplitVBar(I))
Next
End Function
Function DyoSSVbl(SSVbl$) As Variant()
Dim SS: For Each SS In Itr(SplitVBar(SSVbl))
    PushI DyoSSVbl, SyzSS(SS)
Next
End Function

Private Sub Z_DyoVblLy()
Dim VblLy$()
GoSub T1
Exit Sub
T0:
    Erase XX
    X "1 | 2 | 3"
    X "4 | 5 6 |"
    X "| 7 | 8 | 9 | 10 | 11 |"
    X "12"
    VblLy = XX
    Ept = Array(SyzSS("1 2 3"), Sy("4", "5 6", ""), Sy("", "7", "8", "9", "10", "11", ""), Sy("12"))
    GoTo Tst
Exit Sub
T1:
    Erase XX
    X "|lskdf|sdlf|lsdkf"
    X "|lsdf|"
    X "|lskdfj|sdlfk|sdlkfj|sdklf|skldf|"
    X "|sdf"
    VblLy = XX
    Ept = ""
    GoTo Tst
Tst:
    Act = DyoVblLy(VblLy)
    DmpDy CvAv(Act)
'    C
    Return
End Sub

Private Sub Z()
Z_DyoVblLy
MVb_Lin_Vbl:
End Sub
