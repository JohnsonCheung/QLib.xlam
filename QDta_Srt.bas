Attribute VB_Name = "QDta_Srt"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_Srt."
Private Const Asm$ = "QDta"
Dim A_KeyDry()
Dim A_IsDesAy() As Boolean

Private Function SrtDrszFstCol(A As Drs) As Drs
Dim F():      F = FstCol(A)
Dim Ixy&(): Ixy = IxyzSrtAy(F)
  SrtDrszFstCol = DrswRowIxy(A, Ixy)
End Function
Sub Z123()
Z_SrtDrs
End Sub
Sub Z_SrtDrs()
Dim Drs As Drs, Act As Drs, Ept As Drs, SrtByFF$
GoSub T0
Exit Sub
T0:
    SrtByFF = "A B"
    Drs = DrszFF("A B C", DryzSSVbl("4 5 6|1 2 3|2 3 4"))
    Ept = DrszFF("A B C", DryzSSVbl("1 2 3|2 3 4|4 5 6"))
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
Dim Dry():               Dry = SrtDry(A.Dry, ColIxy, Des)
                      SrtDrs = Drs(A.Fny, Dry)
End Function

Function SrtDry(Dry(), SrtColIxy&(), IsDesAy() As Boolean) As Variant()
         A_IsDesAy = IsDesAy
          A_KeyDry = SelDry(Dry, SrtColIxy)
Dim R&():        R = SrtDry__RowIxy
            SrtDry = AywIxy(Dry, R)
End Function

Private Function SrtDrs__IsDesAy(Ay$()) As Boolean()
Dim I: For Each I In Ay
    PushI SrtDrs__IsDesAy, FstChr(I) = "-"
Next
End Function

Private Function SrtDry__RowIxy() As Long()
Dim U&: U = UB(A_KeyDry) ' Always >=1
Dim L&():     L = LngSeqzU(U)          ' Use the LasEle as pivot, so don't include it in L&()
SrtDry__RowIxy = SrtDry__SIxyR(L)
End Function

Private Function SrtDry__TakH(I&, Ixy&()) As Long()
Dim J: For Each J In Ixy
    If Not SrtDry__IsGT(CLng(J), I) Then PushI SrtDry__TakH, J
Next
End Function

Private Function SrtDry__TakL(I&, Ixy&()) As Long()
Dim J: For Each J In Ixy
    If SrtDry__IsGT(CLng(J), I) Then PushI SrtDry__TakL, J
Next
End Function

Private Function SrtDry__Cmp%(Des As Boolean, A, B)
If A = B Then Exit Function
If Des Then
    SrtDry__Cmp = IIf(A > B, 1, -1)
Else
    SrtDry__Cmp = IIf(A < B, 1, -1)
End If
End Function

Private Function SrtDry__IsGT(I1&, I2&) As Boolean
Dim K1: K1 = A_KeyDry(I1)
Dim K2: K2 = A_KeyDry(I2)
Dim A, J&: For Each A In K1
    Dim B:                B = K2(J)
    Dim Des As Boolean: Des = A_IsDesAy(J)
    Select Case SrtDry__Cmp(Des, A, B)
    Case -1: Exit Function
    Case 1: SrtDry__IsGT = True: Exit Function
    End Select
J = J + 1
Next
End Function
Private Function SrtDry__SIxyR(Ixy&()) As Long()
Dim O&()
Dim U&: U = UB(Ixy)
Select Case U
Case -1
Case 0: O = Ixy
Case 1:
    Dim A&: A = Ixy(0)
    Dim B&: B = Ixy(1)
    If SrtDry__IsGT(A, B) Then
        O = LngAp(A, B)
    Else
        O = LngAp(B, A)
    End If
Case Else
    Dim L&(): L = Ixy
    Dim P&:       P = Pop(L)
    Dim P1&():   P1 = SrtDry__TakL(U, L)
    Dim P2&():   P2 = SrtDry__TakH(U, L)
    Dim LAy&(): LAy = SrtDry__SIxyR(P1)
    Dim HAy&(): HAy = SrtDry__SIxyR(P2)
                  O = SrtDry__Add(LAy, P, HAy)
End Select
SrtDry__SIxyR = O
End Function
Private Function SrtDry__Add(LH1&(), P&, LH2&()) As Long()
Dim O&()
Dim L: For Each L In LH1
    PushI O, L
Next
PushI O, P
For Each L In LH2
    PushI O, L
Next
SrtDry__Add = O
End Function

Function SrtDt(A As Dt, Optional SrtByFF$ = "") As Dt
SrtDt = DtzDrs(SrtDrs(DrszDt(A), SrtByFF), A.DtNm)
End Function

Function SrtDryzC(Dry(), C&, Optional IsDes As Boolean) As Variant()
Attribute SrtDryzC.VB_Description = "12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789"
Dim Col(): Col = ColzDry(Dry, C)
Dim Ix&(): Ix = IxyzSrtAy(Col, IsDes)
Dim IFm&, ITo&, IStp%
If IsDes Then
    IFm = 0: ITo = UB(Ix): IStp = 1
Else
    IFm = UB(Ix): ITo = 0: IStp = -1
End If
Dim J&: For J = IFm To ITo Step IStp
   Push SrtDryzC, Dry(Ix(J))
Next
End Function
