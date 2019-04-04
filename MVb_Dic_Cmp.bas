Attribute VB_Name = "MVb_Dic_Cmp"
Option Explicit
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
    O = AyAddAp( _
        FmtExcess(.AExcess, .Nm1), _
        FmtExcess(.BExcess, .Nm2), _
        FmtDif(.ADif, .BDif))
End With
If Not ExlSam Then
    O = AyAdd(O, FmtSam(A.Sam))
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
If A.Count <> B.Count Then Stop
If A.Count = 0 Then Exit Function
Dim O$(), K, S1$, S2$, S(0) As S1S2, Ly$(), KK$
For Each K In A
    KK = K
    S1 = "!" & "Er Dif" & vbCrLf & K & vbCrLf & LinesUnderLin(KK) & vbCrLf & A(K)
    S2 = "!" & "Er Dif" & vbCrLf & K & vbCrLf & LinesUnderLin(KK) & vbCrLf & B(K)
    Set S(0) = S1S2(S1, S2)
    Ly = FmtS1S2Ay(S)
    PushAy O, Ly
Next
FmtDif = O
End Function

Private Function FmtExcess(A As Dictionary, Nm$) As String()
If A.Count = 0 Then Exit Function
Dim K, S1$, S2$, S(0) As S1S2
S2 = "!" & "Er Excess (" & Nm & ")"
For Each K In A.Keys
    S1 = K & vbCrLf & LinesUnderLin(K) & vbCrLf & A(K)
    Set S(0) = S1S2(S1, S2)
    PushAy FmtExcess, FmtS1S2Ay(S)
Next
End Function

Private Function FmtSam(A As Dictionary) As String()
If A.Count = 0 Then Exit Function
Dim O$(), K, S() As S1S2, KK$
For Each K In A.Keys
    KK = K
    PushObj S, S1S2("*Same", K & vbCrLf & LinesUnderLin(KK) & vbCrLf & A(K))
Next
FmtSam = FmtS1S2Ay(S)
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
