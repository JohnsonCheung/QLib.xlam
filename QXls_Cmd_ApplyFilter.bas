Attribute VB_Name = "QXls_Cmd_ApplyFilter"
Option Explicit
Option Compare Text
Public Enum EmOp
   EiNop   ' No operation
   EiPatn
   EiEQ    '=
   EiNE    '!
   EiBET   '%
   EiNBET  '!%
   EiLIS   ':
   EiNLIS  '!:
   EiGE    '>=
   EiGT    '>
   EiLE    '<=
   EiLT    '<
End Enum
Private Sub Z()
Stop
Z_XApply
End Sub

Private Sub Z_XApply()
Dim Rg As Range: Set Rg = CWs.Range("B1")
ApplyFilter Rg
End Sub
Sub ApplyFilter(ByVal Rg As Range)
Dim FCell As Range:      Set FCell = XFCell(Rg) '  # Filter-Cell ! the cell with str-"Filter"
Dim Lo    As ListObject:    Set Lo = XLo(FCell)
If IsNothing(Lo) Then Exit Sub
Dim R2&:                   R2 = Lo.ListColumns(1).Range.Row - 1
Dim C2&:                   C2 = Lo.ListColumns.Count
Dim CriRg As Range: Set CriRg = RgRCRC(FCell, 2, 1, R2, C2)
Select Case Rg.Value
Case "Clear": CriRg.Clear
Case "Apply": XApply Lo, CriRg
End Select
End Sub

Private Function XFCell(Rg As Range) As Range
'Ret :  # Filter-Cell ! the cell with str-"Filter" @@
Dim O As Range
If Rg.Count <> 1 Then Exit Function
Dim C&: C = Rg.Column
Dim V: V = Rg.Value
Select Case True
Case V = "Apply" And C > 1: Set O = RgRC(Rg, 1, 0)
Case V = "Clear" And C > 2: Set O = RgRC(Rg, 1, -1)
Case Else: Exit Function
End Select
If O.Value = "Filter" Then Set XFCell = O
'Insp "QXls_Cmd_ApplyFilter.XFCell", "Inspect", "Oup(XFCell) Rg", "NoFmtr(Range)", "NoFmtr(Range)": Stop
End Function

Private Function XSamCol(A As ListObjects, Cno&) As ListObject()
Dim C As ListObject: For Each C In A
    If C.Range.Column = Cno Then PushObj XSamCol, C
Next
End Function

Private Function XLo(FCell As Range) As ListObject
'Fm FCell :  # Filter-Cell ! the cell with str-"Filter" @@
If IsNothing(FCell) Then Exit Function
Dim Ws     As Worksheet:  Set Ws = WszRg(FCell)
Dim C&:                        C = FCell.Column
Dim SamCol() As ListObject: SamCol = XSamCol(Ws.ListObjects, C)
Dim R&():                      R = XRnoAy(SamCol)
Dim M&:                        M = MinEle(R)
                         Set XLo = XLozWhereR(SamCol, M)
'Insp "QXls_Cmd_ApplyFilter.XLo", "Inspect", "Oup(XLo) FCell", "NoFmtr(ListObject)", "NoFmtr(Range)": Stop
End Function

Private Function XRnoAy(SamCol() As ListObject) As Long()
Dim L: For Each L In Itr(SamCol)
    PushI XRnoAy, CvLo(L).Range.Row
Next
End Function

Private Function XLozWhereR(A() As ListObject, R&) As ListObject
Dim J%: For J = 0 To UB(A)
    If A(J).Range.Row = R Then Set XLozWhereR = A(J): Exit Function
Next
ThwImpossible CSub
End Function

Private Sub XApply(Lo As ListObject, CriRg As Range)
'Fm T : The FilterCell
'Ret  : KSet ! Filter-KSet for each column.  K is the coln V is the vset

'== ClrMsg==============================================================================================================
DltCmtzRg CriRg     '<==
BdrAroundNone CriRg '<==

'-- Fnd Cri & CriBrk ---------------------------------------------------------------------------------------------------
Dim Fny$():             Fny = FnyzLo(Lo)                                             ' from Lo
Dim CriCell As Drs: CriCell = XCriCell(Fny, CriRg)                                   ' F R C CriCell
Dim CriBrk  As Drs:  CriBrk = XCriBrk(CriCell)                                       ' F R C Op V1 V2 Patn T1 IsEr Msg
Dim Cri     As Drs:     Cri = ColEqSel(CriBrk, "IsEr", False, "F Op V1 V2 Patn Sim") ' F Op V1 V2 Patn Sim             ! All-IsEr = false

'== Set Cmt & BdrEr to those cell with CriStr has err ==================================================================
'   Fm  CriBrk
Dim Er       As Drs:       Er = ColEqSel(CriBrk, "IsEr", True, "R C Msg") ' R C Msg
Dim CellAy() As Range: CellAy = RgAy(CriRg, Er)                           ' Cell with er need to BdrEr and set cmt with msg
Dim ErMsg$():           ErMsg = StrCol(Er, "Msg")
Dim OSetCm:                     SetCmtzAy CellAy, ErMsg                   ' <==
Dim OBrkEr:                     BdrErzAy CellAy                           ' <==

'== ShowAllData if there is no Criterior ===============================================================================
If NoReczDrs(Cri) Then
    SetOnFilter Lo
    Lo.AutoFilter.ShowAllData
    Exit Sub
End If

'== Apply Filter with Diff =============================================================================================
Dim CFny$():        CFny = DistColzStr(Cri, "F")  '  # Cri-Fny ! The Fny with cri entered
Dim CCol() As Aset: CCol = VsetAyzLo(Lo, CFny)    '  # Cri-Col ! The col-vset with cri entered
Dim CSel() As Aset: CSel = XCSel(CCol, CFny, Cri) '  # Cri-Sel ! aft the Cri apply, what val in col-vset has been selected
                                                  '            ! [CFny CCol CSet] hav sam nbr of ele
                                                  '            ! ele in @CSel may have no ele, it will consider as warning, need to rpt by
                                                  '            ! by bdr er all the cri cell.
Dim FmRno&: FmRno = Lo.DataBodyRange.Row
Dim ToRno&: ToRno = R2zRg(Lo.DataBodyRange)
Dim EptRny&(): EptRny = XEptRny(FmRno, ToRno, CFny, CSel)     '      # Ept-Rny ! Those row is ept visible
Dim ActRny&(): ActRny = XActRny(Lo)                         '      # Act-Rny ! Those row is visible
Dim ToSetOn&(): '! Rno to set on
Dim ToSetOff&(): '! Rno to set off
Dim Ws As Worksheet: Set Ws = WszLo(Lo)
Dim OupOn:                    XOupOn Ws, ToSetOn       ' <== Turn on  the row
Dim OupOff:                   XOupOff Ws, ToSetOff     ' <== Turn off the row
'== Bdr the cri selecting no record (Ns) (no-sel) ======================================================================
Dim NsFny$(): NsFny = XNsFny(CSel, CFny) '  ! What CFny selecting no rec
If Si(NsFny) > 0 Then
    Dim NsCri    As Drs:     NsCri = DrswIn(CriBrk, "F", NsFny)   ' Cri causing the no sel
    Dim NsExlEr  As Drs:   NsExlEr = ColEq(NsCri, "IsEr", False) ' Exl those already @IsEr
    Dim NsRpt    As Drs:     NsRpt = DrszSel(NsExlEr, "R C")      ' Need to report ns cri
    Dim NsCell() As Range:  NsCell = RgAy(CriRg, NsRpt)
    BdrErzAy NsCell, "Red"
End If
End Sub
Private Sub XOupOn(Ws As Worksheet, ToSetOn&())

End Sub
Private Sub XOupOff(Ws As Worksheet, ToSetOff&())

End Sub
Private Function XEptRny(FmRno&, ToRno&, CFny$(), CSel() As Aset) As Long()

End Function
Private Function XActRny(Lo As ListObject) As Long()

End Function
Private Function XNsFny(CSel() As Aset, CFny$()) As String()
'Fm CSel :  # Cri-Sel ! aft the Cri apply, what val in col-vset has been selected
'        :            ! [CFny CCol CSet] hav sam nbr of ele
'        :            ! ele in @CSel may have no ele, it will consider as warning, need to rpt by
'        :            ! by bdr er all the cri cell.
'Fm CFny :  # Cri-Fny ! The Fny with cri entered
'Ret     :            ! What CFny selecting no rec @@
Dim I, J%: For Each I In CSel
    If CvAset(I).Cnt = 0 Then
        PushI XNsFny, CFny(J)
    End If
Next
'Insp "QXls_Cmd_ApplyFilter.XNsFny", "Inspect", "Oup(XNsFny) CSel CFny", XNsFny, "NoFmtr(() As Aset)", CFny: Stop
End Function

Sub SetCmtzAy(Rg() As Range, Cmt$())
Dim J%: For J = 0 To MinUB(Rg, Cmt)
    SetCmt Rg(J), Cmt(J)
Next
End Sub

Sub Z_SetCmt()
Dim R As Range: Set R = A1zWs(CWs)
SetCmt R, "lskdfjsdlfk"
End Sub

Function HasCmt(R As Range) As Boolean
HasCmt = Not IsNothing(R.Comment)
End Function

Sub SetCmt(R As Range, Cmt$)
If Not HasCmt(R) Then
    R.AddComment.text Cmt
    Exit Sub
End If
Dim C As Comment: Set C = R.Comment
If C.text = Cmt Then Exit Sub
C.text Cmt
End Sub

Sub Z_BdrEr()
BdrEr CWs.Range("C2"), "Red"
End Sub

Sub BdrEr(R As Range, Optional ColrNm$ = "Red")
R.BorderAround xlContinuous, xlMedium, Color:=Colr(ColrNm)
End Sub

Sub BdrErzAy(RgAy() As Range, Optional ColrNm$ = "Red")
Dim R: For Each R In Itr(RgAy)
    BdrEr CvRg(R), ColrNm
Next
End Sub

Function XCSel(CCol() As Aset, CFny$(), Cri As Drs) As Aset()
'Fm CCol :                     # Cri-Col ! The col-vset with cri entered
'Fm CFny :                     # Cri-Fny ! The Fny with cri entered
'Fm Cri  : F Op V1 V2 Patn Sim           ! All-IsEr = false
'Ret     :                     # Cri-Sel ! aft the Cri apply, what val in col-vset has been selected
'        :                               ! [CFny CCol CSet] hav sam nbr of ele
'        :                               ! ele in @CSel may have no ele, it will consider as warning, need to rpt by
'        :                               ! by bdr er all the cri cell. @@
Dim U%: U = UB(CCol)
If U = -1 Then Exit Function
Dim O() As Aset: ReDim O(U)
Dim J%, F: For Each F In CFny
    Dim FCri As Drs: FCri = ColEqE(Cri, "F", F) ' Op V1 V2 Patn T
    Set O(J) = XCSelC(CCol(J), F, FCri)
    J = J + 1
Next
XCSel = O
'Insp "QXls_Cmd_ApplyFilter.XCSel", "Inspect", "Oup(XCSel) CCol CFny Cri", "NoFmtr(Aset())", "NoFmtr(() As Aset)", CFny, LinzDrs(Cri): Stop
End Function

Private Function XCSelC(Col As Aset, F, FCri As Drs) As Aset
Dim Op As EmOp, V1, V2, Patn$, T As EmSimTy, Re As RegExp
Set XCSelC = New Aset
Dim Vy(): Vy = Col.Av
Dim CriDr: For Each CriDr In FCri.Dry
    If Si(Vy) = 0 Then Exit For
    AsgAp CriDr, Op, V1, V2, Patn, T
    Set Re = RegExp(Patn)
    Vy = XCSelVy(Vy, Op, V1, V2, Re, T)
Next
Set XCSelC = AsetzAy(Vy)
End Function

Function VsetzLc(Lc As ListColumn) As Aset
Set VsetzLc = AsetzAy(ColzLc(Lc))
End Function

Function VsetAyzLo(Lo As ListObject, Fny$()) As Aset()
Dim F: For Each F In Itr(Fny)
    Dim Lc As ListColumn: Set Lc = Lo.ListColumns(F)
    Dim Aset As Aset:   Set Aset = VsetzLc(Lc)
    PushObj VsetAyzLo, Aset
Next
End Function

Private Function XCriBrk(CriCell As Drs) As Drs
'Fm CriCell : F R C CriCell
'Ret        : F R C Op V1 V2 Patn T1 IsEr Msg @@
Dim Dr, Dry(): For Each Dr In Itr(CriCell.Dry)
    Dim CriVal: CriVal = Pop(Dr)          ' The cell value with criteria val
    Dim CriAv:   CriAv = XCriAvzC(CriVal) ' Op V1 V2 Patn T1 IsEr Msg        ! brk the criteria str in these 7 values
    PushIAy Dr, CriAv
    PushI Dry, Dr
Next
XCriBrk = DrszFF("F R C Op V1 V2 Patn Sim IsEr Msg", Dry)
'Insp "QXls_Cmd_ApplyFilter.XCriBrk", "Inspect", "Oup(XCriBrk) CriCell", LinzDrs(XCriBrk), LinzDrs(CriCell): Stop
End Function

Private Function XCSelVy(Vy(), Op As EmOp, V1, V2, Re As RegExp, T As EmSimTy) As Variant()
Dim V: For Each V In Vy
    If XIsSel(V, Op, V1, V2, Re, T) Then
        PushI XCSelVy, V
    End If
Next
End Function

Private Function XIsSel(V, Op As EmOp, V1, V2, Re As RegExp, T1 As EmSimTy) As Boolean
On Error GoTo X
Dim O As Boolean
Select Case True
Case Op = EiBET:  O = IsBet(V, V1, V2)
Case Op = EiNBET: O = Not IsBet(V, V1, V2)
Case Op = EiGE:   O = V >= V1
Case Op = EiGT:   O = V > V1
Case Op = EiLE:   O = V <= V1
Case Op = EiLT:   O = V < V1
Case Op = EiLIS:  O = HasEle(V1, V)
Case Op = EiNLIS: O = Not HasEle(V1, V)
Case Op = EiNE:   O = V <> V1
Case Op = EiEQ:   O = V = V1
Case Op = EiPatn: O = Re.Test(V)
Case Else: Thw CSub, "Op error"
End Select
XIsSel = O
Exit Function
X: Dim E$: E = Err.Description
   Inf CSub, "Runtime er", "Er V-To-Be-Select Op V1 V2 Patn T1", E, V, Op, V1, V2, Re.Pattern, T1
End Function

Private Function XCriCell(Fny$(), CriRg As Range) As Drs
'Fm Fny : from Lo
'Ret    : F R C CriCell @@
Dim Sq(): Sq = CriRg.Value
Dim CriCol()
Dim C%, F, Dry(): For Each F In Fny
    C = C + 1
    CriCol = ColzSq(Sq, C)
    Dim R%: R = 0
    Dim CriCell: For Each CriCell In CriCol
        R = R + 1
        If Not IsEmpty(CriCell) Then
            PushI Dry, Array(F, R, C, CriCell)
        End If
    Next
Next
XCriCell = DrszFF("F R C CriCell", Dry)
'Insp "QXls_Cmd_ApplyFilter.XCriCell", "Inspect", "Oup(XCriCell) Fny CriRg", LinzDrs(XCriCell), Fny, "NoFmtr(Range)": Stop
End Function

Private Function XCriAv(Op As EmOp, V1, Optional V2, Optional Patn$, Optional IsEr As Boolean, Optional Msg$) As Variant()
XCriAv = Array(Op, V1, V2, Patn, SimTy(V1), IsEr, Msg)
End Function

Private Function XCriAvzC(CriVal)
'Fm CriVal : The cell value with criteria val @@
Dim V: V = CriVal
Dim T As EmSimTy: T = SimTy(V)
Dim O()
Select Case T
Case EiYes, EiDte, EiNbr: O = XCriAv(EiEQ, V)
Case EiStr:               O = XCriAvzStr(CStr(V))
End Select
XCriAvzC = O
'Insp "QXls_Cmd_ApplyFilter.XCriAv", "Inspect", "Oup(XCriAv) CriVal", XCriAv, CriVal: Stop
End Function

Private Function XShfOp(OStr$) As EmOp
'   EiPatn
'   EiEQ    '=
'   EiNE    '!
'   EiBET   '%
'   EiNBET  '!%
'   EiLIS   ':
'   EiNLIS  '!:
'   EiGE    '>=
'   EiGT    '>
'   EiLE    '<=
'   EiLT    '<
Dim F$: F = FstChr(OStr)
Dim O As EmOp
Select Case F
Case "=": O = EiEQ: GoTo X1
Case "!"
    If SndChr(OStr) = ":" Then
        O = EiNLIS
        GoTo X2
    End If
    O = EiNE
    GoTo X1
Case "%": O = EiBET: GoTo X1
Case ":": O = EiLIS: GoTo X1
Case ">"
    If SndChr(OStr) = "=" Then
        O = EiGE
        GoTo X2
    End If
    O = EiGT
    GoTo X1
Case "<"
    If SndChr(OStr) = "=" Then
        O = EiLE
        GoTo X2
    End If
    O = EiLT
    GoTo X1
Case Else
    XShfOp = EiPatn
End Select
Exit Function
X1:
    XShfOp = O: OStr = RmvFstChr(OStr)
    Exit Function
X2:
    XShfOp = O: OStr = Mid(OStr, 3)
End Function

Private Function XCriAvzBet(Op As EmOp, IsStr As Boolean, S$) As Variant()
Dim V1$, V2$: AsgTRst S, V1, V2
Dim O()
Select Case True
Case IsStr:                       O = XCriAv(Op, V1, V2)
Case IsNbrzS(V1) And IsNbrzS(V2): O = XCriAv(Op, CDbl(V1), CDbl(V2))
Case IsDtezS(V1) And IsNbrzS(V2): O = XCriAv(Op, CDate(V1), CDate(V2))
Case Else:                        O = XCriAv(Op, V1, V2)
End Select
If V1 > V2 Then Inf CSub, "V1 cannot > V2", "Ty V1 V2", TypeName(V1), V1, V2: Exit Function
XCriAvzBet = O
End Function

Private Function XCriAvzLis(Op As EmOp, IsStr As Boolean, S$) As Variant()
Dim Sy$(): Sy = SyzSS(S)
If IsDteSy(Sy) Then XCriAvzLis = XCriAv(Op, DteAyzSy(Sy)): Exit Function
If IsDblSy(Sy) Then XCriAvzLis = XCriAv(Op, DblAyzSy(Sy)): Exit Function
XCriAvzLis = XCriAv(Op, Sy)
End Function

Private Function XCriAvzSng(Op As EmOp, IsStr As Boolean, S$) As Variant()
Dim A$: A = Trim(S)
If IsStr Then
    XCriAvzSng = XCriAv(Op, S)
    Exit Function
End If
If IsDtezS(S) Then
    Dim D As Date: D = A
    XCriAvzSng = XCriAv(Op, D)
    Exit Function
End If

If IsNbrzS(S) Then
    Dim Dbl#: Dbl = A
    XCriAvzSng = XCriAv(Op, Dbl)
    Exit Function
End If
XCriAvzSng = XCriAv(Op, A)
End Function
Private Function XCriAvzStr(S$) As Variant()
Dim Op As EmOp: Op = XShfOp(S)
If S = "" Then
    Inf CSub, "S should be blank"
    Exit Function
End If
Dim L$: L = S
Dim IsStr As Boolean: If Op <> EiPatn Then If ShfPfx(L, "'") Then IsStr = True
L = Trim(L)
Select Case Op
Case EiBET, EiNBET:                      XCriAvzStr = XCriAvzBet(Op, IsStr, L)
Case EiLE, EiLT, EiGT, EiGE, EiNE, EiEQ: XCriAvzStr = XCriAvzSng(Op, IsStr, L)
Case EiLIS, EiNLIS:                      XCriAvzStr = XCriAvzLis(Op, IsStr, L)
Case EiPatn:                             XCriAvzStr = XCriAv(Op, Empty, Patn:=L)
Case Else:                               XCriAvzStr = XCriAvzEr("Invalid EmOp from CellVal")
End Select
End Function
Private Function XCriAvzEr(Msg$) As Variant()
XCriAvzEr = XCriAv(EiNop, "", IsEr:=True, Msg:=Msg)
End Function
Private Sub XApplyzDif(KSetzDif As Dictionary, Lo As ListObject)
'Fm Fld : Fld [V]  'Idx is the filter index
'Ret    : For each F in Fld, apply the filter
Dim K: For Each K In KSetzDif.Keys
    XApplyzLc Lo, K, CvAset(KSetzDif(K))
Next
End Sub

Private Sub XApplyzLc(Lo As ListObject, F, S As Aset)
Dim Fld%:  Fld = Lo.ListColumns(F).Index
Dim Sel(): Sel = S.Av
SetOnFilter Lo
Stop
If Si(Sel) > 0 Then 'If no data is selected
Lo.Range.AutoFilter Fld, Sel, xlFilterValues
End If
Stop
End Sub

Private Function NotUse_XAsk$()
Dim T$: T = "Filter"
Erase XX
X "[Yes] = Apply"
X "[No] = Clear"
Dim M$: M = JnCrLf(XX)
Dim S As VbMsgBoxStyle: S = vbYesNoCancel + vbQuestion + vbDefaultButton1
Dim R As VbMsgBoxResult: R = MsgBox(M, S, T)
Dim O$
Select Case True
Case R = vbYes: O = "Apply"
Case R = vbNo: O = "Clear"
End Select
NotUse_XAsk = O
End Function


