Attribute VB_Name = "MDta_Dry"
Option Explicit
Const CMod$ = "MDta_Dry."
Function CntDrywGT1(CntDry()) As Variant()
Dim Dr
For Each Dr In CntDry
    If Dr(1) > 1 Then PushI CntDrywGT1, Dr
Next
End Function

Function DrywColInAy(A(), ColIx%, InAy) As Variant()
Const CSub$ = CMod & "DrywColInAy"
If Not IsArray(InAy) Then Thw CSub, "[InAy] is not Array, but [TypeName]", InAy, TypeName(InAy)
If Sz(InAy) = 0 Then DrywColInAy = A: Exit Function
Dim Dr
For Each Dr In Itr(A)
    If HasEle(InAy, Dr(ColIx)) Then PushI DrywColInAy, Dr
Next
End Function
Sub C3DryDo3(C3Dry(), Do3$)
If Sz(C3Dry) = 0 Then Exit Sub
Dim Dr
For Each Dr In C3Dry
    Run Do3, Dr(0), Dr(1), Dr(2)
Next
End Sub

Sub C4DryDo4(C4Dry(), Do4$)
If Sz(C4Dry) = 0 Then Exit Sub
Dim Dr
For Each Dr In C4Dry
    Run Do4, Dr(0), Dr(1), Dr(2), Dr(3)
Next
End Sub
Function DotNyDry(A()) As String()
Dim Dr
For Each Dr In Itr(A)
    PushI DotNyDry, JnDot(Dr)
Next
End Function
Function DryDotNy(DotNy$()) As Variant()
If Sz(DotNy) = 0 Then Exit Function
Dim O(), I
For Each I In DotNy
   With Brk1(I, ".")
       Push O, Sy(.S1, .S2)
   End With
Next
DryDotNy = O
End Function

Private Sub ZZ_FmtA()
DmpAy FmtDry(SampDry1)
End Sub

Function DrywColHasDup(A(), C) As Variant()
DrywColHasDup = AywIxAy(A, IxAyzDup(ColzDry(A, C)))
End Function
Private Sub Z_DryzJnFldKK()
Dim Dry(), KKIxAy, JnFldIx, Sep$
Sep = " "
Dry = Array(Array(1, 2, 3, 4, "A"), Array(1, 2, 3, 6, "B"), Array(1, 2, 2, 8, "C"), Array(1, 2, 2, 12, "DD"), _
Array(2, 3, 1, 1, "x"), Array(12, 3), Array(12, 3, 1, 2, "XX"))
Ept = Array()
KKIxAy = Array(0, 1, 2)
JnFldIx = 4
GoSub Tst
Exit Sub
Tst:
    Act = DryzJnFldKK(Dry, KKIxAy, JnFldIx, Sep)
    BrwDry CvAv(Act)
    StopNE
    Return
End Sub
Function DrsJnFldKKFld(DRs As DRs, KK, JnFld, Optional Sep$ = " ") As DRs

End Function

Function DryzJnFldKK(Dry(), KKIxAy, JnFldIx, Optional Sep$ = " ") As Variant()
Dim IxAy: IxAy = KKIxAy: PushI IxAy, JnFldIx
Dim N%: N = Sz(IxAy)
DryzJnFldKK = DryJnFldNFld(DrySelIxAy(Dry, IxAy), N)
End Function

Function RowIxzDryDr&(Dry(), Dr)
Dim N%: N = Sz(Dr)
Dim Ix&, D
For Each D In Itr(Dry)
    If IsEqAy(AywFstNEle(D, N), Dr) Then
        RowIxzDryDr = Ix
        Exit Function
    End If
    Ix = Ix + 1
Next
RowIxzDryDr = -1
End Function
Function DryJnFldNFld(Dry(), FstNFld%, Optional Sep$ = " ") As Variant()
Dim U%: U = FstNFld - 1
Dim UK%: UK = U - 1
Dim O(), Dr
For Each Dr In Itr(Dry)
    If U <> UB(Dr) Then
        ReDim Preserve Dr(U)
    End If
    Dim Ix: Ix = RowIxzDryDr(O, AywFstNEle(Dr, UK))
    If Ix = -1 Then
        PushI O, Dr
    Else
        O(Ix)(U) = Appd(O(Ix)(U), Sep) & Dr(U)
    End If
Next
DryJnFldNFld = O
End Function
Function DryzSslAy(SslAy) As Variant()
Dim L
For Each L In Itr(SslAy)
    PushI DryzSslAy, SySsl(L)
Next
End Function

Function CntDiczDry(A(), C) As Dictionary
Set CntDiczDry = CntDic(ColzDry(A, C))
End Function

Function SeqCntDiczDry(A(), C) As Dictionary
Set SeqCntDiczDry = SeqCntDicvAy(ColzDry(A, C))
End Function

Function SqlTyzDryC$(A(), C)
SqlTyzDryC = SqlTyzAv(ColzDry(A, C))
End Function
Function SqlTyzAv$(Av())
Dim O As VbVarType, V, T As VbVarType
For Each V In Av
    T = VarType(V)
    If T = vbString Then
        If Len(V) > 255 Then SqlTyzAv = "Memo": Exit Function
    End If
    O = MaxVbTy(O, T)
Next
End Function
Function SqlTyzVbTy$(A As VbVarType, Optional IsMem As Boolean)
Select Case A
Case vbEmpty: SqlTyzVbTy = "Text(255)"
Case vbBoolean: SqlTyzVbTy = "YesNo"
Case vbByte: SqlTyzVbTy = "Byte"
Case vbInteger: SqlTyzVbTy = "Short"
Case vbLong: SqlTyzVbTy = "Long"
Case vbDouble: SqlTyzVbTy = "Double"
Case vbSingle: SqlTyzVbTy = "Single"
Case vbCurrency: SqlTyzVbTy = "Currency"
Case vbDate: SqlTyzVbTy = "Date"
Case vbString: SqlTyzVbTy = IIf(IsMem, "Memo", "Text(255)")
Case Else: Stop
End Select
End Function

Function DryInsAv(A(), Av())

End Function



Function IsBrkDryIxC(A(), DrIx&, BrkColIx) As Boolean
If Sz(A) = 0 Then Exit Function
If DrIx = 0 Then Exit Function
If DrIx = UB(A) Then Exit Function
If A(DrIx)(BrkColIx) = A(DrIx - 1)(BrkColIx) Then Exit Function
IsBrkDryIxC = True
End Function

Function NColDry%(A)
Dim O%, Dr
For Each Dr In Itr(A)
    O = Max(O, Sz(Dr))
Next
NColDry = O
End Function



Function NRowDryCEv&(A(), C, Ev)
Dim J&, O&, Dr
For Each Dr In Itr(A)
   If Dr(C) = Ev Then O = O + 1
Next
NRowDryCEv = O
End Function



Function DrywCEv(A(), C, Ev) As Variant()
Dim O()
Dim Dr
For Each Dr In Itr(A)
   If Dr(C) = Ev Then PushI DrywCEv, Dr
Next
End Function

Function DrywCCNe(A, C1, C2) As Variant()
Dim Dr
For Each Dr In A
    If Dr(C1) <> Dr(C2) Then PushI DrywCCNe, Dr
Next
End Function

Sub ThwNEDry(A(), B())
If Not IsEqDry(A, B) Then Stop
End Sub


Function DrywColEq(A, C%, V) As Variant()
Dim Dr
For Each Dr In A
    If Dr(C) = V Then PushI DrywColEq, Dr
Next
End Function

Function DrywCGt(A, C%, GtV) As Variant()
Dim Dr
For Each Dr In Itr(A)
    If Dr(C) > GtV Then PushI DrywCGt, Dr
Next
End Function

Function DrywDup(A(), C) As Variant()
Dim Dup, Dr, O()
Dup = AywDup(ColzDry(A, C))
For Each Dr In Itr(A)
    If HasEle(Dup, Dr(C)) Then Push O, Dr
Next
DrywDup = O
End Function

Function DrywIxAyzy(A, IxAy, EqVy) As Variant()
Dim Dr
For Each Dr In A
    If IsEqAy(AywIxAy(Dr, IxAy), EqVy) Then PushI DrywIxAyzy, Dr
Next
End Function

Function DistSyzDry(A(), C) As String()
DistSyzDry = AywDist(SyzDry(A, C))
End Function

Function DryzSqCol(Sq, ColIxAy) As Variant()
Dim R&
For R = 1 To UBound(Sq, 1)
    PushI DryzSqCol, DrzSqr(Sq, R)
Next
End Function

Function DryzSq(Sq) As Variant()
Dim R&
For R = 1 To UBound(Sq, 1)
    PushI DryzSq, DrzSqr(Sq, R)
Next
End Function


Function HasDr(Dry(), Dr) As Boolean
Dim IDr
For Each IDr In Itr(Dry)
    If IsEqAy(IDr, Dr) Then HasDr = True: Exit Function
Next
End Function

