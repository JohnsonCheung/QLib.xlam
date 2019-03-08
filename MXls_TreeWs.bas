Attribute VB_Name = "MXls_TreeWs"
Option Explicit
Private LasHom$
Sub Change(Target As Range)
If Not IsA1(Target) Then Exit Sub
Dim Ws As Worksheet: Set Ws = WszRg(Target)
If Not IsActionWs(Ws) Then Exit Sub
Dim mA1 As Range: Set mA1 = A1(Ws)
EnsA1 mA1
If Not IsPth(mA1.Value) Then Exit Sub
Dim Hom$: Hom = mA1.Value
If LasHom = Hom Then Exit Sub
ShwEntzHom Hom
'ShwFstHomFdr Ws
End Sub

Sub SelectionChange(Target As Range)
If Target.Row = 1 Then Exit Sub
Static WIP As Boolean
If WIP Then Debug.Print "MTreeWs.SelectionChange: WIP": Exit Sub
WIP = True
Dim Ws As Worksheet: Set Ws = WszRg(Target)
Stop
EnsA1 A1(Ws)
If Not IsAction(Ws) Then Exit Sub
ShwCurCol Target
ShwNxtCol Target
WIP = False
End Sub
Private Sub ShwEntzHom(Hom$)
ShwEnt Hom, 1
End Sub
Private Sub ShwEnt(Pth$, Cno%)
Dim FnAy$(), FdrAy$()
AsgEnt FdrAy, FnAy, Pth
ShwEntzPut Cno, FdrAy, FnAy
End Sub
Private Sub ShwEntzPut(Cno%, FdrAy$(), FnAy$())

End Sub
Private Sub ShwFstHomFdr(Hom$)

End Sub
Private Sub ShwCurCol(Cur As Range)
ShwCurEnt Cur
End Sub


Private Sub ShwNxtEnt()

End Sub

Private Sub ShwCurEnt(Cur As Range)
ClrCurCol Cur
Dim SubPthAy$(), FnAy$()
AsgEnt SubPthAy, FnAy, PthzCur(Cur)
PutCurEnt Cur, SubPthAy, FnAy
MgeCurSubPthCol Sz(SubPthAy)
MgeCurFnCol Sz(SubPthAy), Sz(FnAy)
End Sub

Private Function PthzCur$(Cur As Range)
PthzCur = PthEnsSfx(A1zRg(Cur).Value)
End Function
Private Sub PutCurEnt(Cur As Range, SubPthAy$(), FnAy$())
EntRg(Cur, Sz(SubPthAy) + Sz(FnAy)).Value = SqzAyV(AyAdd(SubPthAy, FnAy))
End Sub
Private Function EntRg(Cur As Range, EntCnt%) As Range
Dim Ws As Worksheet: Set Ws = WszRg(Cur)
Set EntRg = WsCRR(Ws, Cur.Column, 2, EntCnt + 1)
End Function

Private Sub ClrCurCol(Cur As Range)
Dim Ws As Worksheet: Set Ws = WszRg(Cur)
WsCRR(Ws, Cur.Column, 2, LasCno(Ws)).Delete
End Sub

Private Function CurColCC() As Range

End Function
Private Function MgeCurSubPthCol(SubPthSz&)

End Function
Private Function MgeCurFnCol(SubPthSz&, FnSz&)

End Function

Private Sub ShwNxtCol(Cur As Range)
ShwRow Cur
End Sub
Private Sub ShwRow(Cur As Range)
Dim Ws As Worksheet: Set Ws = WszRg(Cur)
Dim R%: R = MaxR(Ws)
Dim LasR&: LasR = LasRno(Ws)
WsRR(Ws, 1, R).Hidden = False
WsRR(Ws, R + 1, LasR).Hidden = True
End Sub
Private Function MaxR%(Ws As Worksheet)
Dim J%, O%
For J% = 1 To MaxC(Ws)
    O = Max(O, WsRC(Ws, 2, J).End(xlDown).Row - 1)
Next
MaxR = O
End Function
Private Function MaxC%(Ws As Worksheet)
MaxC = CnozBefFstHid(Ws)
End Function
Private Sub EnsA1(A1 As Range)
If IsActionA1(A1) Then Exit Sub
A1.Value = "Please enter a valid path here"
Clear WszRg(A1)
End Sub
Private Sub Clear(Ws As Worksheet)
A1zWs(Ws).Activate
DltColFm Ws, 2
DltRowFm Ws, 2
HidColFm Ws, 2
HidRowFm Ws, 2
WsC(Ws, 1).AutoFit
End Sub
Private Function IsAction(Ws As Worksheet) As Boolean
IsAction = True
If IsActionWs(Ws) Then Exit Function
If IsActionA1(A1(Ws)) Then Exit Function
IsAction = False
End Function
Private Function IsActionWs(Ws As Worksheet) As Boolean
IsActionWs = Ws.Name = "TreeWs"
End Function

Private Function IsActionA1(A1 As Range) As Boolean
Dim V: V = A1.Value
If Not IsStr(V) Then Exit Function
IsActionA1 = HasPth(V)
End Function
