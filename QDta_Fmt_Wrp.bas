Attribute VB_Name = "QDta_Fmt_Wrp"
Option Explicit
Private Const CMod$ = "MDta_Fmt_Wrp."
Private Const Asm$ = "QDta"
Function WrpDrNRow%(WrpDr())
Dim Col, R%, M%
For Each Col In Itr(WrpDr)
    M = Si(Col)
    If M > R Then R = M
Next
WrpDrNRow = R
End Function

Function SyWrpPad(Sy$(), W%) As String() ' Each Itm of Sy[Sy] is padded to line with WdtzAy(Sy).  return all padded lines as String()
Dim O$(), X, I%
ReDim O(0)
For Each X In Itr(Sy)
    If Len(O(I)) + Len(X) < W Then
        O(I) = O(I) & X
    Else
        PushI O, X
        I = I + 1
    End If
Next
SyWrpPad = O
End Function

Function WrpDrPad(WrpDr, W%()) As Variant() _
'Some Cell in WrpDr may be an array, Wrping each element to cell if their width can fit its W%(?)
Dim J%, Cell, O()
O = WrpDr
For Each Cell In Itr(O)
    If IsArray(Cell) Then
        Stop
'        O(J) = SyWrpPad(Cell, W(J))
    End If
    J = J + 1
Next
WrpDrPad = O
End Function

Function SqzWrpDr(WrpDr()) As Variant()
Dim O(), R%, C%, NCol%, NRow%, Cell, Col, NColi%
NCol = Si(WrpDr)
NRow = WrpDrNRow(WrpDr)
ReDim O(1 To NRow, 1 To NCol)
C = 0
For Each Col In WrpDr
    C = C + 1
    If IsArray(Col) Then
        NColi = Si(Col)
        For R = 1 To NRow
            If R <= NColi Then
                O(R, C) = Col(R - 1)
            Else
                O(R, C) = ""
            End If
        Next
    Else
        O(1, C) = Col
        For R = 2 To NRow
            O(R, C) = ""
        Next
    End If
Next
SqzWrpDr = O
End Function

Function WdtzWrpgDry(WrpgDry(), WrpWdt%) As Integer() _
'WrpgDry is dry having 1 or more wrpCol, which mean need Wrpping.
'WrpWdt is for wrpCol. _
'WrpCol is col with each cell being array.
'if maxWdt of array-ele of wrpCol has wdt > WrpWdt, use that wdt
'otherwise use WrpWdt
If Si(WrpgDry) = 0 Then Exit Function
Dim J&, Col()
For J = 0 To NColzDry(WrpgDry) - 1
    Col = ColzDry(WrpgDry, J)
    If IsArray(Col(0)) Then
'        Push WdtzWrpgDry, WdtzAy(AyFlat(Col))
    Else
'        Push WdtzWrpgDry, WdtzAy(Col)
    End If
Next
End Function


Function WrpCellzDr(Dr(), ColWdt%()) As String()
Dim X
For Each X In Itr(Dr)
'    PushIAy WrpCellzDr, WrpCell(X, ColWdt)
Next
End Function

Function FmtDrWrp(WrpDr, W%()) As String() _
'Each Itm of WrpDr may be an array.  So a AlignLzDrW return Ly not string.
Dim Dr(): Dr = WrpDrPad(WrpDr, W)
Dim Sq(): Sq = SqzWrpDr(Dr)
Dim Sq1(): Sq1 = AlignSq(Sq(), W)
Dim Ly$(): Ly = LyzSq(Sq1)
PushIAy FmtDrWrp, Ly
End Function

Function DryWrpCell(A(), Optional WrpWdt% = 40) As String() _
'WrpWdt is for wrp-col.  If maxWdt of an ele of wrp-col > WrpWdt, use the maxWdt
Dim W%(), Dr, A1(), M$()
W = WdtzWrpgDry(A, WrpWdt)
For Each Dr In Itr(A)
    M = FmtDrWrp(Dr, W)
    PushIAy DryWrpCell, M
Next
End Function


Function AlignSq(Sq(), WdtAy%()) As Variant()
If UBound(Sq(), 2) <> Si(WdtAy) Then Stop
Dim C%, R%, W%, O
O = Sq
For C = 1 To UBound(Sq(), 2) - 1 ' The last column no need to align
    W = WdtAy(C - 1)
    For R = 1 To UBound(Sq(), 1)
        O(R, C) = AlignL(CStr(Sq(R, C)), W)
    Next
Next
AlignSq = O
End Function
