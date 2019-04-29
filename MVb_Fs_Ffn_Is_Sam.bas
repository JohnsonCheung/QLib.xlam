Attribute VB_Name = "MVb_Fs_Ffn_Is_Sam"
Option Explicit
Function IsDifFfn(A$, B$, Optional UseNotEq As Boolean) As Boolean
If UseNotEq Then
    IsDifFfn = Not IsEqFfn(A, B)
Else
    IsDifFfn = Not IsSamFfn(A, B)
End If
End Function
Function IsEqFfn(A$, B$, Optional UseEq As Boolean) As Boolean
ThwIfFfnNotExist A, CSub, "Fst File"
If A = B Then Thw CSub, "Fil A and B are eq name", "A", A
If Not HasFfn(B) Then Exit Function
If Not IsSamSzFfn(A, B) Then Exit Function
Dim J&, F1%, F2%
F1 = FnoRnd128(A)
F2 = FnoRnd128(B)
For J = 1 To NBlk(SizFfn(A), 128)
    If FnoBlk(F1, J) <> FnoBlk(F2, J) Then
        Close #F1, F2
        Exit Function
    End If
Next
Close #F1, F2
IsEqFfn = True
End Function

Function IsSamFfn(A$, B$) As Boolean
If DtezFfn(A) <> DtezFfn(B) Then Exit Function
If Not IsSamSzFfn(A, B) Then Exit Function
IsSamFfn = True
End Function

Function IsSamSzFfn(A$, B$) As Boolean
IsSamSzFfn = SizFfn(A) = SizFfn(B)
End Function

Function MsgSamFfn(A, B, Si&, Tim$, Optional Msg$) As String()
Dim O$()
Push O, "File 1   : " & A
Push O, "File 2   : " & B
Push O, "File Size: " & Si
Push O, "File Time: " & Tim
Push O, "File 1 and 2 have same size and time"
If Msg <> "" Then Push O, Msg
MsgSamFfn = O
End Function

Private Sub Z_FfnBlk()
Dim T$, S$, A$
S = "sllksdfj lsdkjf skldfj skldfj lk;asjdf lksjdf lsdkfjsdkflj "
T = TmpFt
WrtStr S, T
Debug.Assert SizFfn(T) = Len(S)
A = FfnBlk(T, 1)
Debug.Assert A = Left(S, 128)
End Sub

Function FnoBlk$(Fno%, IBlk)
Dim A As String * 128
Get #Fno, IBlk, A
FnoBlk = A
End Function

Function FfnBlk$(Ffn$, IBlk)
Dim F%: F = FnoRnd(Ffn$, 128)
FfnBlk = FnoBlk(F, IBlk)
Close #F
End Function

