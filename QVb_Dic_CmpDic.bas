Attribute VB_Name = "QVb_Dic_CmpDic"
Option Explicit
Private Const CMod$ = "MVb_Dic_Cmp."
Private Const Asm$ = "QVb"
Private Type DicCmp
    Nm1 As String
    Nm2 As String
    AExcess As Dictionary
    BExcess As Dictionary
    ADif As Dictionary
    BDif As Dictionary
    Sam As Dictionary
End Type
Function FmtCmpDic(A As Dictionary, B As Dictionary, Optional Nm1$ = "Fst", Optional Nm2$ = "Snd") As String()
FmtCmpDic = FmtDicCmp(DicCmp(A, B, Nm1, Nm2))
End Function

Function FmtDicCmp(A As DicCmp, Optional ExlSam As Boolean) As String()
Dim O$()
With A
    O = AddAyAp( _
        FmtExcess(.AExcess, .Nm1), _
        FmtExcess(.BExcess, .Nm2), _
        FmtDif(.ADif, .BDif))
End With
If Not ExlSam Then
    O = AddAy(O, FmtSam(A.Sam))
End If
FmtDicCmp = O
End Function

Function DicCmp(A As Dictionary, B As Dictionary, Nm1$, Nm2$) As DicCmp
With DicCmp
    .Nm1 = Nm1
    .Nm2 = Nm2
    Set .AExcess = DicMinus(A, B)
    Set .BExcess = DicMinus(B, A)
    Set .Sam = DicSamKV(A, B)
    AsgADifBDif A, B, .ADif, .BDif
End With
End Function

Sub BrwCmpDicAB(A As Dictionary, B As Dictionary, Optional Nm1$ = "Fst", Optional Nm2$ = "Snd")
BrwAy FmtCmpDic(A, B, Nm1, Nm2)
End Sub

Function DicSamKV(A As Dictionary, B As Dictionary) As Dictionary
Set DicSamKV = New Dictionary
If A.Count = 0 Or B.Count = 0 Then Exit Function
Dim K
For Each K In A.Keys
    If B.Exists(K) Then
        If A(K) = B(K) Then
            DicSamKV.Add K, A(K)
        End If
    End If
Next
End Function

Private Sub AsgADifBDif(A As Dictionary, B As Dictionary, _
    OADif As Dictionary, OBDif As Dictionary)
Dim K
Set OADif = New Dictionary
Set OBDif = New Dictionary
For Each K In A.Keys
    If B.Exists(K) Then
        If A(K) <> B(K) Then
            OADif.Add K, A(K)
            OBDif.Add K, B(K)
        End If
    End If
Next
End Sub

Private Function FmtDif(A As Dictionary, B As Dictionary) As String()
If A.Count <> B.Count Then Thw CSub, "Dic A & B should have same size", "Dic-A-Si Dic-B-Si", A.Count, B.Count
If A.Count = 0 Then Exit Function
Dim O$(), K, S1$, S2$, S As S1S2s, KK$
For Each K In A
    KK = K
    S1 = "!" & "Er Dif" & vbCrLf & K & vbCrLf & UnderLinzLines(KK) & vbCrLf & A(K)
    S2 = "!" & "Er Dif" & vbCrLf & K & vbCrLf & UnderLinzLines(KK) & vbCrLf & B(K)
    PushS1S2 S, S1S2(S1, S2)
Next
FmtDif = FmtS1S2s(S, Nm1:="", Nm2:="")
End Function

Private Function FmtExcess(A As Dictionary, Nm$) As String()
If A.Count = 0 Then Exit Function
Dim K, S1$, S2$, S As S1S2s
S2 = "!" & "Er Excess (" & Nm & ")"
For Each K In A.Keys
    S1 = UnderLinzLines(CStr(K))
    S2 = A(K)
    PushS1S2 S, S1S2(S1, S2)
Next
PushAy FmtExcess, FmtS1S2s(S, Nm1:="Exccess", Nm2:=Nm)
End Function

Private Function FmtSam(A As Dictionary) As String()
If A.Count = 0 Then Exit Function
Dim O$(), K, S As S1S2s, KK$
For Each K In A.Keys
    KK = K
    PushS1S2 S, S1S2("*Same", K & vbCrLf & UnderLinzLines(KK) & vbCrLf & A(K))
Next
FmtSam = FmtS1S2s(S)
End Function

Private Sub Z_BrwCmpDicAB()
Dim A As Dictionary, B As Dictionary
Set A = DiczVbl("X AA|A BBB|A Lines1|A Line3|B Line1|B line2|B line3..")
Set B = DiczVbl("X AA|C Line|D Line1|D line2|B Line1|B line2|B line3|B Line4")
BrwCmpDicAB A, B
End Sub

Private Sub Z()
Z_BrwCmpDicAB
End Sub
