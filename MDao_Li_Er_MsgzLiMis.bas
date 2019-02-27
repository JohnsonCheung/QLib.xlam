Attribute VB_Name = "MDao_Li_Er_MsgzLiMis"
Option Explicit

Function MsgzLiMis(A As LiMis) As String()
MsgzLiMis = AyAddAp(MsgzMisFfnAset(A.MisFfn), MisMsgTbl(A.MisTbl), MisMsgCol(A.MisCol), MisMsgTy(A.MisTy))
End Function

Private Function MisMsgTbl(A() As LiMisTbl) As String()
Dim J%
For J = 0 To UB(A)
    PushIAy MisMsgTbl, A(J).MisMsg
Next
End Function

Private Function MisMsgCol(A() As LiMisCol) As String()
Dim J%
For J = 0 To UB(A)
    PushIAy MisMsgCol, A(J).MisMsg
Next
End Function

Private Property Get MisMsgTy(A() As LiMisTy) As String()
Dim J%, O$()
For J = 0 To UB(A)
    PushIAy O, A(J).MisMsg
Next
End Property


Private Sub Z_ChkCol()
Dim WPth$, LiPm As LiPm, Act$(), Ept$()
'LiPm = ShpCstPm.Li
GoSub Tst
Exit Sub
Tst:
'    CpyFfnAyToPthIfDif ExistFfnAy(FxAyLiFil(LiPm.Fil)), WPth
'    Act = ChkCol(LiPm, WPth)
    D Act
    Stop
    If Not IsEqAy(Act, Ept) Then Stop
    Return
End Sub



