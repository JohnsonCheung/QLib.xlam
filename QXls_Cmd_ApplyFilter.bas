Attribute VB_Name = "QXls_Cmd_ApplyFilter"
Public Enum EmOp
   EiPatn
   EiNE
   EiBET
   EiNBET
   EiLIS
   EiNLIS
   EiGE
   EiGT
   EiLE
   EiLT
End Enum
Private Type Criteria
    Ty As EmSimTy 'True for Num, False For Str
    Op As EmOp
    V1 As Variant
    V2 As Variant
End Type
Private Type Criterias: N As Integer: Ay() As Criteria: End Type

Sub ApplyFilter(ByVal Rg As Range)
Dim FCell As Range:  Set FCell = XFCell(Rg)
Dim Lo As ListObject:   Set Lo = XLo(FCell)
If IsNothing(Lo) Then Exit Sub

Dim R2&:                   R2 = Lo.ListColumns(1).Range.Row - 1
Dim C2&:                   C2 = Lo.ListColumns.Count
Dim CriRg As Range: Set CriRg = RgRCRC(T, 2, 1, R2, C2)
Select Case Rg.Value
Case "Clear": CriRg.Clear
Case "Apply": XApply Lo, CriRg
End Select
End Sub

Private Function XFCell(Rg As Range) As Range
Dim O As Range
If Rg.Count <> 1 Then Exit Function
Dim C&: C = Rg.Column
Select Case Rg.Value
Case "Apply" And C > 1: Set O = RgRC(Rg, 1, 0)
Case "Clear" And C > 2: Set O = RgRC(Rg, 1, -1)
End Select
If O.Value = "Filter" Then Set XFCell = O
End Function

Private Function XSamCol(A As ListObjects, Cno&) As ListObject()
Dim C As ListObject: For Each C In A
    If C.Range.Column = Cno Then PushObj XSamCol, C
Next
End Function

Private Function XLo(FCell As Range) As ListObject
Dim Ws     As Worksheet:  Set Ws = WszRg(FCell)
Dim C&:                        C = FCell.Column
Dim SamCol As ListObject: SamCol = XSamCol(Ws.ListObjects, C)
Dim R&():                      R = XRnoAyzSamCol(SamCol)
Dim M&:                        M = MinEle(R)
                         Set XLo = XLozWhereR(SamCol, M)
End Function

Private Function XRnoAyzSamCol(SamCol() As ListObject) As Long()
Dim L: For Each L In SamCol
    PushI XRnoAyzSamCol, CvLo(L).Range.Row
Next
End Function

Private Function XLozWhereR(A() As ListObject, R&) As ListObject
For J = 0 To UB(A)
    If A(J).Range.Row = R Then Set XLozWhereR = A(J): Exit Function
Next
ThwImpossible CSub
End Function

Private Sub XApply(Lo As ListObject, CriRg As Range)
Dim Act As Dictionary: Set Act = KSetzLoFilter(Lo)
Dim Ept As Dictionary: Set Ept = XEpt(Lo, CriRg)           ' Ept-Filter-KSet
Dim Dif As Dictionary: Set Dif = DifKSet(Ept, Act) ' KSet !
XApplyzDif Dif, Lo
End Sub

Private Function XCriVset(Lo As ListObject, C, Cris As Criterias) As Aset
Dim Lc As ListColumn: Set Lc = L.ListColumns(C)
Dim Col(): Col = AywDist(ColzLc(Lc))
Set XCriVset = XSelVset(Col, Cris)
End Function
Private Function XCris(CriCol()) As Criterias
Dim CriAy(): CriAy = AyeEmpEle(CriCol)
Dim N%: N = Si(CriAy)
If N = 0 Then Exit Function
XCris.N = N
ReDim XCris.Ay(N - 1)
Dim J%: For J = 0 To N - 1
    XCris.Ay(J) = XCri(CriAy(J))
Next
End Function

Private Function XSelVset(Col(), C As Criterias) As Aset
Set XSetVset = New Aset
Dim Ay() As Criteria: Ay = C.Ay
For Each V In Col
    Dim IsSel As Boolean: IsSel = True
    Dim J%: For J = 0 To C.N - 1
        If Not XIsSel(V, Ay(J)) Then IsSel = False: Exit For
    Next
    If IsSel Then XSetVset.PushItm V
Next
End Function

Private Function XIsSel(V, C As Criteria) As Boolean
Dim V1: If Cri.Op <> EiPatn Then V1 = Cri.V1
Dim Op As EmOp: Op = Cri.Op
Select Case True
Case Op = EiBET:  O = IsBet(V, V1, Cri.V2)
Case Op = EiNBET: O = Not IsBet(V, V1, Cri.V2)
Case Op = EiGE:   O = V >= V1
Case Op = EiGT:   O = V > V1
Case Op = EiLE:   O = V <= V1
Case Op = EiLT:   O = V < V1
Case Op = EiLIS:  O = HasEle(V1, V)
Case Op = EiNLIS: O = Not HasEle(V1, V)
Case Op = EiNE:   O = V <> V1
Case Op = EiPatn: O = CvRe(Cri.V1).Test(V)
Case Else: Thw CSub, "Op error"
End Select
XIsSel = O
End Function

Private Function XCrisAy(Lo As ListObject, CriRg As Range) As Criterias()
Dim NCol%: NCol = Lo.ListColumns.Count
Dim Sq(): Sq = CriRg.Value
Dim O() As Criterias
ReDim O(NCol - 1)
For C = 1 To NCol
    O(C) = XCris(ColzSq(Sq, C))
Next
XCrisAy = O
End Function

Private Function XCri(Ty As EmSimTy, CriCellVal) As Criteria
SimTyzCol
End Function
Private Function XEpt(Lo As ListObject, CriRg As Range) As Dictionary
'Fm T : The FilterCell
'Ret  : KSet ! Filter-KSet for each column.  K is the coln V is the vset
                          Set XEpt = New Dictionary
Dim CrisAy() As Criterias: CrisAy = XCrisAy(CriRg)
Dim Fny$():                    Fny = FnyzLo(Lo)
Dim F, J%: For Each F In Fny
    Dim Cris As Criteria: Cris = CrisAy(J)
    Dim V As Aset:        Set V = XCriVset(Lo, F, Cris)
                                  If V.Cnt >= 0 Then
                                     XEpt.Add F, V
                                  End If
                              J = J + 1
Next
End Function

Private Sub XApplyzDif(DifKSet As Dictionary, Lo As ListObject)
'Fm Fld : Fld [V]  'Idx is the filter index
'Ret    : For each F in Fld, apply the filter
Dim K: For Each K In DifKSet.Keys
    Dim Lc As ListColumn: Set Lc = Lo.ListColumns(K)
    Dim Fld%: Fld = Lc.Index
    Dim S As Aset: Set Sel = DifKSet(K)
    Dim Sel(): S.Av
    R.AutoFilter Fld, Sel, xlFilterValues
Next
End Sub

Private Function XAsk$()
Dim T$: T = "Filter"
Erase XX
X "[Yes] = Apply"
X "[No] = Clear"
Dim M$: M = JnCrLf(XX)
Dim S As VbMsgBoxStyle: S = vbYesNoCancel + vbQuestion + vbDefaultButton1
Dim R As VbMsgBoxResult: R = MsgBox(M, S, T)
Select Case True
Case R = vbYes: XAsk = "Apply"
Case R = vbNo: XAsk = "Clear"
End Select
End Function


