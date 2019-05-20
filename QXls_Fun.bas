Attribute VB_Name = "QXls_Fun"
Attribute VB_Description = "aaa"
Option Compare Text
Option Explicit
Private Const CMod$ = "MXls_Fun."
Private Const Asm$ = "QXls"

Sub FillAtV(At As Range, Ay)
FillSq Sqv(Ay), At
End Sub

Sub FillLc(Lo As ListObject, ColNm$, Ay)
If Lo.DataBodyRange.Rows.Count <> Si(Ay) Then Thw CSub, "Lo-NRow <> Si(Ay)", "Lo-NRow ,Si(Ay)", NRowzLo(Lo), Si(Ay)
Dim At As Range, C As ListColumn, R As Range
'DmpAy FnyzLo(Lo)
'Stop
Set C = Lo.ListColumns(ColNm)
Set R = C.DataBodyRange
Set At = R.Cells(1, 1)
FillAtV At, Ay
End Sub
Sub FillSq(Sq(), At As Range)
ResiRg(At, Sq).Value = Sq
End Sub
Sub FillAtH(Ay, At As Range)
FillSq Sqh(Ay), At
End Sub

Function WszAyab(A, B, Optional N1$ = "Ay1", Optional N2$ = "Ay2") As Worksheet
Set WszAyab = WszDrs(DrszAyab(A, B, N1, N2))
End Function

Function WszDic(Dic As Dictionary, Optional InclDicValOptTy As Boolean) As Worksheet
Set WszDic = WszDrs(DrszDic(Dic, InclDicValOptTy))
End Function

Function WszDt(A As Dt) As Worksheet
Dim O As Worksheet
Set O = NewWs(A.DtNm)
LozDrs DrszDt(A), A1zWs(O)
WszDt = O
End Function

Function NyzFml(Fml$) As String()
NyzFml = NyzMacro(Fml)
End Function

Function WszLy(Ly$(), Optional Wsn$ = "Sheet1") As Worksheet
Set WszLy = WszRg(RgzAyV(Ly, A1zWs(NewWs(Wsn))))
End Function

Property Get MaxWsCol&()
Static C&, Y As Boolean
If Not Y Then
    Y = True
    C = IIf(Xls.Version = "16.0", 16384, 255)
End If
MaxWsCol = C
End Property

Property Get MaxWsRow&()
Static R&, Y As Boolean
If Not Y Then
    Y = True
    R = IIf(Xls.Version = "16.0", 1048576, 65535)
End If
MaxWsRow = R
End Property

Function SqHzN(N%) As Variant()
Dim O()
ReDim O(1 To 1, 1 To N)
SqHzN = O
End Function

Function SqVzN(N%) As Variant()
Dim O(), J%
ReDim O(1 To N, 1 To 1)
SqVzN = O
End Function

Function N_ZerFill$(N, NDig&)
N_ZerFill = Format(N, String(NDig, "0"))
End Function

Function WszS1S2s(A As S1S2s, Optional Nm1$ = "S1", Optional Nm2$ = "S2") As Worksheet
Set WszS1S2s = WszSq(SqzS1S2s(A, Nm1, Nm2))
End Function

Private Sub Z_AyabWs()
GoTo ZZ
Dim A, B
ZZ:
    A = SyzSS("A B C D E")
    B = SyzSS("1 2 3 4 5")
    ShwWs WszAyab(A, B)
End Sub

Private Sub Z_WbFbOupTbl()
GoTo ZZ
ZZ:
    Dim W As Workbook
    'Set W = WbFbOupTbl(WFb)
    'ShwWb W
    Stop
    'W.Close False
    Set W = Nothing
End Sub
