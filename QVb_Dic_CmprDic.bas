Attribute VB_Name = "QVb_Dic_CmprDic"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Dic_Cmp."
Private Const Asm$ = "QVb"
Private Type CmpgDic
    Nm1 As String
    Nm2 As String
    AExcess As Dictionary
    BExcess As Dictionary
    ADif As Dictionary
    BDif As Dictionary
    Sam As Dictionary
End Type
Function FmtCmpgDic(A As Dictionary, B As Dictionary, Optional Nm1$ = "Fst", Optional Nm2$ = "Snd") As String()
FmtCmpgDic = FmtCmpgDiczCmpg(CmpgDic(A, B, Nm1, Nm2))
End Function

Function FmtCmpgDiczCmpg(A As CmpgDic, Optional ExlSam As Boolean) As String()
Dim O$()
With A
    O = AyzAddAp( _
        FmtExcess(.AExcess, .Nm1), _
        FmtExcess(.BExcess, .Nm2), _
        FmtDif(.ADif, .BDif))
End With
If Not ExlSam Then
    O = AyzAdd(O, FmtSam(A.Sam))
End If
FmtCmpgDiczCmpg = O
End Function

Function CmpgDic(A As Dictionary, B As Dictionary, Nm1$, Nm2$) As CmpgDic
With CmpgDic
    .Nm1 = Nm1
    .Nm2 = Nm2
    Set .AExcess = MinusDic(A, B)
    Set .BExcess = MinusDic(B, A)
    Set .Sam = DicSamKV(A, B)
    AsgADifBDif A, B, .ADif, .BDif
End With
End Function

Sub BrwCmpgDicAB(A As Dictionary, B As Dictionary, Optional Nm1$ = "Fst", Optional Nm2$ = "Snd")
BrwAy FmtCmpgDic(A, B, Nm1, Nm2)
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
FmtDif = FmtS1S2s(S, N1:="", N2:="")
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
PushAy FmtExcess, FmtS1S2s(S, N1:="Exccess", N2:=Nm)
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

Private Sub Z_BrwCmpgDicAB()
Dim A As Dictionary, B As Dictionary
Set A = DiczVbl("X AA|A BBB|A Lines1|A Line3|B Line1|B line2|B line3..")
Set B = DiczVbl("X AA|C Line|D Line1|D line2|B Line1|B line2|B line3|B Line4")
BrwCmpgDicAB A, B
End Sub

Private Sub Z()
Z_BrwCmpgDicAB
End Sub
