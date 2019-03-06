VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Rel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const CMod$ = "Rel."
Public Nm$
Private Dic As New Dictionary    ' Key is Par, Val is Dic of chd

Friend Function Init(RelLy$()) As Rel
Dim O As New Rel, L
For Each L In Itr(RelLy)
    O.PushRelLin L
Next
Set Init = O
End Function
Sub PushParChd(P, C)
Dim S As Aset
If Dic.Exists(P) Then
    Set S = Dic(P)
    S.PushItm C
Else
    Set S = New Aset
    S.PushItm C
    Dic.Add P, S
End If
End Sub
Sub PushRelLin(A)
Dim Ay$(), P$, C
Ay = SySsl(A)
If Sz(Ay) = 0 Then Exit Sub
P = AyShf(Ay)
For Each C In Itr(Ay)
    PushParChd P, C
Next
End Sub
Sub Brw()
BrwAy Fmt
End Sub
Function Clone() As Rel
Dim O As New Rel
Set Clone = O.Init(Fmt)
End Function

Sub Dmp()
D Fmt
End Sub

Property Get Fmt() As String()
Dim K
For Each K In Dic.Keys
    PushI Fmt, ParLin(K)
Next
End Property

Function IsEq(A As Rel) As Boolean
Stop '
'If Not IsEqItr(A.Rel.Keys, B.Rel.Keys) Then Exit Function
'Dim K
'For Each K In Rel_ParAset(A)
'    If Not Aset_IsEq(A.Rel(K), B.Rel(K)) Then Exit Function
'Next
'Rel_IsEq = True
End Function

Sub ThwIfNE(A As Rel, Optional Msg$ = "Two rel are diff", Optional ANm$ = "Rel-B")
Const CSub$ = CMod & "ThwIfNE"
If IsEq(A) Then Exit Sub
Dim O$()
PushI O, Msg
PushI O, FmtQQ("?-ParCnt(?) / ?-ParCnt(?)", Nm, NPar, ANm, A.NPar)
PushI O, Nm & " --------------------"
PushIAy O, Fmt
PushI O, ANm & " --------------------"
PushIAy O, A.Fmt
ThwErMsg O, CSub, "Two rel not eq"
End Sub

Sub ThwNotVdt()
Const CSub$ = CMod & "ThwNotVdt"
Dim I
For Each I In Dic.Values
    If Not IsAset(I) Then
        Thw CSub, "Given Rel is not a valid due to the chd of K is not Aset", "Rel K [TypeName of K's Chd]", Fmt, I, TypeName(Dic(I))
    End If
Next
End Sub

Property Get NItm&()
NItm = Itms.Cnt
End Property

Function IsLeaf(Itm) As Boolean
IsLeaf = True
If IsNoChdPar(Itm) Then Exit Function
If Not IsPar(Itm) Then Exit Function
IsLeaf = False
End Function

Function IsNoChdPar(Itm) As Boolean
If Not IsPar(Itm) Then Exit Function
IsNoChdPar = ParChd(Itm).IsEmp
End Function

Function IsPar(Itm) As Boolean
IsPar = Dic.Exists(Itm)
End Function

Function Itms() As Aset
Dim O As New Aset, K
O.PushItr Dic.Keys
For Each K In Dic.Keys
    O.PushAset Dic(K)
Next
Set Itms = O
End Function

Function InDpdOrdItms() As Aset
Const CSub$ = CMod & "InDpdOrdItms"
'Return itms in Rel in dependant order. Throw er if there is cyclic
'Example: A B C D
'         C D E
'         E X
'Return: B D X E C A
Dim O As New Aset, J%, M As Rel, Leaves As Aset
Set M = Clone
Do
    J = J + 1: If J > 1000 Then Thw CSub, "looping to much"
    Set Leaves = M.Leaf
    If Leaves.IsEmp Then
        If M.NPar > 0 Then
            Thw CSub, "Cyclic relation is found so far.  No leaves but there is remaining Rel", _
            "Turn-Cnt [Orginal rel] [Dpd itm found] [Remaining relation not solved]", _
            J, Fmt, O.Lin, M.Fmt
        End If
        Set InDpdOrdItms = O
        Exit Function
    End If
    O.PushAset Leaves
    M.RmvAllLeaf
    O.PushAset M.NoChdPar
    M.RmvNoChdPar
Loop
Set InDpdOrdItms = O
End Function

Function Par() As Aset
Set Par = AsetzAy(Dic.Keys)
End Function

Function Leaf() As Aset
Dim Itm, O As New Aset
For Each Itm In Itms.Itms
    If IsLeaf(Itm) Then O.PushItm Itm
Next
Set Leaf = O
End Function

Function NoChdPar() As Aset
Dim O As New Aset, P
For Each P In Par.Itms
    If ParIsNoChd(P) Then O.PushItm P
Next
Set NoChdPar = O
End Function

Sub HasPar_XAss(Par, Fun$)
If IsPar(Par) Then Exit Sub
Thw Fun, "Given Par is not a parent", "Rel Par", Fmt, Par: Stop
End Sub

Property Get NPar&()
NPar = Dic.Count
End Property

Function ParHasChd(P, C) As Boolean
If IsPar(P) Then Exit Function
ParHasChd = ParChd(P).Has(C)
End Function
Function ParChd(P) As Aset
If Dic.Exists(P) Then Set ParChd = Dic(P)
End Function
Function ParIsNoChd(P) As Boolean
If Not IsPar(P) Then Exit Function
ParIsNoChd = CvAset(Dic(P)).IsEmp
End Function

Function ParLin$(P)
If Not IsPar(P) Then Exit Function
ParLin = P & " " & ParChd(P).Lin
End Function
Function ParRmvChdAy&(P, ChdAy())
If Not IsPar(P) Then Exit Function
Dim C, O&
For Each C In Itr(ChdAy)
    If ParChd(P).RmvItm(C) Then
        O = O + 1
    End If
Next
ParRmvChdAy = O
End Function

Function ParRmvChdItm(P, C) As Boolean
Dim X As Aset
If ParHasChd(P, C) Then
    ParRmvChdItm = ParChd(P).RmvItm(C)
End If
End Function

Function RmvAllLeaf&()
Dim P, LeafAy(), O&
LeafAy = Leaf.Av
For Each P In Dic.Keys
    If ParRmvChdAy(P, LeafAy) Then
        O = O + 1
    End If
Next
RmvAllLeaf = O
End Function

Function RmvNoChdPar&()
Dim P, O&
For Each P In NoChdPar.Itms
    Dic.Remove P
    O = O + 1
Next
RmvNoChdPar = O
End Function

Function RmvPar(P) As Boolean
If IsPar(P) Then
    Dic.Remove P
    RmvPar = True
    Exit Function
End If
End Function

Property Get SampRel() As Rel
Set SampRel = RelVbl("B C D | D E | X")
End Property

Friend Sub Z_Itms()
Dim Act As Aset, Ept As Aset, A As Rel
Set Ept = AsetzSsl("A B C D E")
Set A = RelVbl("A B C | B D E | C D")
GoSub Tst
Exit Sub
Tst:
    Set Act = A.Itms
    C
    Return
End Sub

Friend Sub Z_InDpdOrdItms()
Dim Act As Aset, Ept As Aset
Dim R As Rel
GoSub T1
'GoSub T2
Exit Sub
T1:
    Set Ept = AsetzSsl("C E X D B")
    Set R = RelVbl("B C D | D E | X")
    GoSub Tst
    Return
'
T2:
    Dim X$()
    PushI X, "MVb"
    PushI X, "MIde MVb MXls MAcs"
    PushI X, "MXls MVb"
    PushI X, "MDao MVb MDta"
    PushI X, "MAdo MVb"
    PushI X, "MAdoX MVb"
    PushI X, "MApp  MVb"
    PushI X, "MDta  MVb"
    PushI X, "MTp   MVb"
    PushI X, "MSql  MVb"
    PushI X, "AStkShpCst MVb MXls MAcs"
    PushI X, "MAcs  MVb MXls"
    Set R = Rel(X)
    Set Ept = AsetzSsl("MVb MIde MXls MDao MAdo MAdoX MApp MDta MTp MSql AStkShpCst MAcs ")
    GoSub Tst
    Return
Tst:
    Set Act = R.InDpdOrdItms
    If Not Act.IsEq(Ept) Then Stop
    Return
End Sub

Private Sub ZZ()
Dim A As Variant
Dim B$()
Dim C$
Dim D As Rel

CvRel A
IsRel A
End Sub

Friend Sub Z()
Me.Z_InDpdOrdItms
Me.Z_Itms
End Sub


