Attribute VB_Name = "QVb_Lin_Vbl"
Option Explicit
Private Const CMod$ = "MVb_Lin_Vbl."
Private Const Asm$ = "QVb"

Function DryzTLiny(TLiny$()) As Variant()
Dim I
For Each I In Itr(TLiny)
    PushI DryzTLiny, TermAy(I)
Next
End Function

Function DryzVblLy(A$()) As Variant()
Dim I
For Each I In Itr(A)
    PushI DryzVblLy, AyTrim(SplitVBar(I))
Next
End Function

Private Sub Z_DryzVblLy()
Dim VblLy$()
Push VblLy, "1 | 2 | 3"
Push VblLy, "4 | 5 6 |"
Push VblLy, "| 7 | 8 | 9 | 10 | 11 |"
Push VblLy, "12"
Ept = Array(SyzSS("1 2 3"), Sy("4", "5 6", ""), Sy("", "7", "8", "9", "10", "11", ""), Sy("12"))
GoSub Tst
Exit Sub
Tst:
    Act = DryzVblLy(VblLy)
'    Ass IsEqDry(CvAy(Act), CvAy(Ept))
    Return
End Sub

Private Sub ZZ_DryzVblLy()
Dim VblLy$()
Push VblLy, "|lskdf|sdlf|lsdkf"
Push VblLy, "|lsdf|"
Push VblLy, "|lskdfj|sdlfk|sdlkfj|sdklf|skldf|"
Push VblLy, "|sdf"
DmpDry DryzVblLy(VblLy)
End Sub

Private Sub ZZ()
Z_DryzVblLy
MVb_Lin_Vbl:
End Sub
