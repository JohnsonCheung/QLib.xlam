VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Aset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const CMod$ = "Aset."
Private Aset As New Dictionary
Property Get TermLin$()
Dim I, O$()
For Each I In Itms
    PushI O, QuoteSqIf(CStr(I))
Next
TermLin = JnSpc(O)
End Property

Property Get Cnt&()
Cnt = Aset.Count
End Property

Sub Dmp()
D Aset.Keys
End Sub
Sub Vc()
Brw UseVc:=True
End Sub
Sub Brw(Optional Fnn$, Optional UseVc As Boolean)
MVb_Fun.Brw Aset.Keys, DftStr(Fnn, "Aset"), UseVc
End Sub
Function Srt() As Aset
Set Srt = AsetzAy(QSrt1(Itms))
End Function
Function AddAset(A As Aset) As Aset
Dim O As New Aset
O.PushAset Aset
O.PushAset A
Set AddAset = O
End Function

Function RmvItm(Itm) As Aset
If Has(Itm) Then Aset.Remove Itm
Set RmvItm = Me
End Function

Sub PushItm(Itm)
If Not Has(Itm) Then Aset.Add Itm, Empty
End Sub
Sub PushAy(A)
Dim I
For Each I In Itr(A)
    PushItm I
Next
End Sub
Sub PushItr(Itr, Optional NoBlankStr As Boolean)
Dim I
If NoBlankStr Then
    For Each I In Itr
        If I <> "" Then
            PushItm I
        End If
    Next
Else
    For Each I In Itr
        PushItm I
    Next
End If
End Sub

Function Clone() As Aset
Dim O As New Aset
O.PushItr Aset.Keys
Set Clone = O
End Function

Function Minus(B As Aset) As Aset
Dim O As Aset, I
Set O = EmpAset
For Each I In Itms
    If Not B.Has(I) Then O.PushItm I
Next
Set Minus = O
End Function
Function Has(Itm) As Boolean
Has = Aset.Exists(Itm)
End Function

Function IsEq(B As Aset) As Boolean
If Cnt <> B.Cnt Then Exit Function
Dim K
For Each K In Itms
    If Not B.Has(K) Then Exit Function
Next
IsEq = True
End Function


Function IsEmp() As Boolean
IsEmp = Aset.Count = 0
End Function
Function Av() As Variant()
Av = AvzItr(Itms)
End Function

Function IsInOrdEq(B As Aset) As Boolean
IsInOrdEq = IsEqAy(Me.Av, B.Av)
End Function
Function FstItm()
Const CSub$ = CMod & "FstItm"
If IsEmp Then Thw CSub, "Given Aset is empty"
Dim I
For Each I In Itms
    Asg I, FstItm
Next
End Function
Function AbcDic() As Dictionary 'AbcDic means the keys is comming from Aset the value is starting from A, B, C
Dim O As New Dictionary, J%, K
For Each K In Me.Itms
    O.Add K, Chr(65 + J%)
    J = J + 1
Next
Set AbcDic = O
End Function
Function Itms()
Itms = Aset.Keys
End Function

Function Lin$()
Lin = JnSpc(Aset.Keys)
End Function

Sub PushAset(A As Aset)
PushItr A.Itms
End Sub

Function Sy() As String()
Sy = SyzAy(Aset.Keys)
End Function

Private Sub ZZ()
Dim A As Variant
Dim B As Aset
Dim C$

End Sub

Private Sub Z()
End Sub

