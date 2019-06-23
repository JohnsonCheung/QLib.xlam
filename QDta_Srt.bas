Attribute VB_Name = "QDta_Srt"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_Srt."
Private Const Asm$ = "QDta"
Dim A_KeyDy()
Dim A_IsDesAy() As Boolean

Private Function SrtDrszFstCol(A As Drs) As Drs
Dim F():      F = FstCol(A)
Dim Ixy&(): Ixy = IxyzSrtAy(F)
  SrtDrszFstCol = DwRxy(A, Ixy)
End Function

Private Sub Z_SrtDrs()
Dim Drs As Drs, Act As Drs, Ept As Drs, SrtByFF$
GoSub T0
Exit Sub
T0:
    SrtByFF = "A B"
    Drs = DrszFF("A B C", DyoSSVbl("4 5 6|1 2 3|2 3 4"))
    Ept = DrszFF("A B C", DyoSSVbl("1 2 3|2 3 4|4 5 6"))
    GoTo Tst
Tst:
    Act = SrtDrs(Drs, SrtByFF)
    If Not IsEqDrs(Act, Ept) Then Stop
    Return
End Sub

Function SrtDrs(A As Drs, Optional SrtByFF$ = "") As Drs
'Fm SrtByFF : If SrtByFF is blank use fst col. @
If NoReczDrs(A) Then SrtDrs = A: Exit Function
If SrtByFF = "" Then
    SrtDrs = SrtDrszFstCol(A)
    Exit Function
End If
Dim Ay$():                Ay = Ny(SrtByFF)           ' Each ele may have - as pfx, which mean descending
Dim Fny$():              Fny = RmvPfxzAy(Ay, "-")
Dim ColIxy&():        ColIxy = Ixy(A.Fny, Fny)
Dim Des() As Boolean:    Des = SrtDrs__IsDesAy(Ay)
Dim Dy():               Dy = SrtDy(A.Dy, ColIxy, Des)
                      SrtDrs = Drs(A.Fny, Dy)
End Function

Function SrtDy(Dy(), SrtColIxy&(), IsDesAy() As Boolean) As Variant()
         A_IsDesAy = IsDesAy
          A_KeyDy = DyoSel(Dy, SrtColIxy)
Dim R&():        R = SrtDy__Rxy
            SrtDy = AwIxy(Dy, R)
End Function

Private Function SrtDrs__IsDesAy(Ay$()) As Boolean()
Dim I: For Each I In Ay
    PushI SrtDrs__IsDesAy, FstChr(I) = "-"
Next
End Function

Private Function SrtDy__Rxy() As Long()
Dim U&: U = UB(A_KeyDy) ' Always >=1
Dim L&():     L = LngSeqzU(U)          ' Use the LasEle as pivot, so don't include it in L&()
SrtDy__Rxy = SrtDy__SIxyR(L)
End Function

Private Function SrtDy__TakH(I&, Ixy&()) As Long()
Dim J: For Each J In Ixy
    If Not SrtDy__IsGT(CLng(J), I) Then PushI SrtDy__TakH, J
Next
End Function

Private Function SrtDy__TakL(I&, Ixy&()) As Long()
Dim J: For Each J In Ixy
    If SrtDy__IsGT(CLng(J), I) Then PushI SrtDy__TakL, J
Next
End Function

Private Function SrtDy__Cmp%(Des As Boolean, A, B)
If A = B Then Exit Function
If Des Then
    SrtDy__Cmp = IIf(A > B, 1, -1)
Else
    SrtDy__Cmp = IIf(A < B, 1, -1)
End If
End Function

Private Function SrtDy__IsGT(I1&, I2&) As Boolean
Dim K1: K1 = A_KeyDy(I1)
Dim K2: K2 = A_KeyDy(I2)
Dim A, J&: For Each A In K1
    Dim B:                B = K2(J)
    Dim Des As Boolean: Des = A_IsDesAy(J)
    Select Case SrtDy__Cmp(Des, A, B)
    Case -1: Exit Function
    Case 1: SrtDy__IsGT = True: Exit Function
    End Select
J = J + 1
Next
End Function
Private Function SrtDy__SIxyR(Ixy&()) As Long()
Dim O&()
Dim U&: U = UB(Ixy)
Select Case U
Case -1
Case 0: O = Ixy
Case 1:
    Dim A&: A = Ixy(0)
    Dim B&: B = Ixy(1)
    If SrtDy__IsGT(A, B) Then
        O = LngAp(A, B)
    Else
        O = LngAp(B, A)
    End If
Case Else
    Dim L&(): L = Ixy
    Dim P&:       P = Pop(L)
    Dim P1&():   P1 = SrtDy__TakL(U, L)
    Dim P2&():   P2 = SrtDy__TakH(U, L)
    Dim LAy&(): LAy = SrtDy__SIxyR(P1)
    Dim HAy&(): HAy = SrtDy__SIxyR(P2)
                  O = SrtDy__Add(LAy, P, HAy)
End Select
SrtDy__SIxyR = O
End Function
Private Function SrtDy__Add(LH1&(), P&, LH2&()) As Long()
Dim O&()
Dim L: For Each L In LH1
    PushI O, L
Next
PushI O, P
For Each L In LH2
    PushI O, L
Next
SrtDy__Add = O
End Function

Function SrtDt(A As Dt, Optional SrtByFF$ = "") As Dt
SrtDt = DtzDrs(SrtDrs(DrszDt(A), SrtByFF), A.DtNm)
End Function

Function SrtDyoC(Dy(), C&, Optional IsDes As Boolean) As Variant()
Attribute SrtDyoC.VB_Description = "12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789"
Dim Col(): Col = ColzDy(Dy, C)
Dim Ix&(): Ix = IxyzSrtAy(Col, IsDes)
Dim IFm&, ITo&, IStp%
If IsDes Then
    IFm = 0: ITo = UB(Ix): IStp = 1
Else
    IFm = UB(Ix): ITo = 0: IStp = -1
End If
Dim J&: For J = IFm To ITo Step IStp
   Push SrtDyoC, Dy(Ix(J))
Next
End Function
