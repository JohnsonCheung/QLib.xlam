Attribute VB_Name = "MXls_Fun"
Attribute VB_Description = "aaa"
Option Explicit

Sub PutAyColAt(A, At As Range)
Dim Sq()
Sq = SqzAyV(A)
RgzResz(At, Sq).Value = Sq
End Sub

Sub PutAyRgzLc(A, Lo As ListObject, ColNm$)
Dim At As Range, C As ListColumn, R As Range
'DmpAy FnyzLo(Lo)
'Stop
Set C = Lo.ListColumns(ColNm)
Set R = C.DataBodyRange
Set At = R.Cells(1, 1)
PutAyColAt A, At
End Sub

Sub PutAyRowAt(Ay, At As Range)
Dim Sq()
Sq = SqzAyH(Ay)
RgzResz(At, Sq).Value = Sq
End Sub

Function AyabWs(A, B, Optional N1$ = "Ay1", Optional N2$ = "Ay2", Optional LoNm$ = "AyAB") As Worksheet
Dim N&, AtA1 As Range, R As Range
N = Si(A)
If N <> Si(B) Then Stop
Set AtA1 = NewA1

PutAyRowAt Array(N1, N2), AtA1
PutAyColAt A, AtA1.Range("A2")
PutAyColAt B, AtA1.Range("B2")
'LozRg RgRCRC(AtA1, 1, 1, N + 1, 2)
Set AyabWs = AtA1.Parent
End Function



Function NewWsDic(Dic As Dictionary, Optional InclDicValOptTy As Boolean) As Worksheet
Set NewWsDic = WszDrs(DrszDic(Dic, InclDicValOptTy))
End Function
Function NewWsVisDic(A As Dictionary, Optional InclDicValOptTy As Boolean) As Worksheet
Set NewWsVisDic = WsVis(NewWsDic(A, InclDicValOptTy))
End Function

Function NewWsDt(A As Dt, Optional Vis As Boolean) As Worksheet
Dim O As Worksheet
Set O = NewWs(A.DtNm)
LozDrs DrszDt(A), A1zWs(O)
Set NewWsDt = O
If Vis Then WsVis O
End Function

Function NyFml(A$) As String()
NyFml = NyzMacro(A)
End Function

Sub SetLcTotLnk(A As ListColumn)
Dim R1 As Range, R2 As Range, R As Range, Ws As Worksheet
Set R = A.DataBodyRange
Set Ws = WszRg(R)
Set R1 = RgRC(R, 0, 1)
Set R2 = RgRC(R, R.Rows.Count + 1, 1)
Ws.Hyperlinks.Add Anchor:=R1, Address:="", SubAddress:=R2.Address
Ws.Hyperlinks.Add Anchor:=R2, Address:="", SubAddress:=R1.Address
R1.Font.ThemeColor = xlThemeColorDark1
End Sub

Function LyWs(Ly$(), Vis As Boolean) As Worksheet
Dim O As Worksheet: Set O = NewWs()
'AyRgV Ly, A1zWs(O)
Set LyWs = O
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

Function SqHBar(N%) As Variant()
Dim O()
ReDim O(1 To 1, 1 To N)
SqHBar = O
End Function

Function SqVbar(N%) As Variant()
Dim O(), J%
ReDim O(1 To N, 1 To 1)
SqVbar = O
End Function

Function N_ZerFill$(N, NDig%)
N_ZerFill = Format(N, String(NDig, "0"))
End Function

Function WszS1S2Ay(A() As S1S2, Optional Nm1$ = "S1", Optional Nm2$ = "S2") As Worksheet
Set WszS1S2Ay = WszSq(SqzS1S2Ay(A, Nm1, Nm2))
End Function

Private Sub Z_AyabWs()
GoTo ZZ
Dim A, B
ZZ:
    A = SySsl("A B C D E")
    B = SySsl("1 2 3 4 5")
    WsVis AyabWs(A, B)
End Sub

Private Sub Z_WbFbOupTbl()
GoTo ZZ
ZZ:
    Dim W As Workbook
    'Set W = WbFbOupTbl(WFb)
    'WbVis W
    Stop
    'W.Close False
    Set W = Nothing
End Sub
