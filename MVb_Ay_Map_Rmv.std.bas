Attribute VB_Name = "MVb_Ay_Map_Rmv"
Option Explicit
Private Sub Y(S$, X$)
PushI XX, RplQ(S, X)
End Sub
Function AyRmvFstChr(A) As String()
Dim I
For Each I In Itr(A)
    PushI AyRmvFstChr, RmvFstChr(I)
Next
End Function

Function AyRmvFstNonLetter(A) As String() 'Gen:AyXXX
Dim I
For Each I In Itr(A)
    PushI AyRmvFstNonLetter, RmvFstNonLetter(I)
Next
End Function
Function AyRmvLasChr(A) As String()
'Gen:AyFor RmvLasChr
Dim I
For Each I In Itr(A)
    PushI AyRmvLasChr, RmvLasChr(I)
Next
End Function

Function AyRmvPfx(A, Pfx) As String()
If Sz(A) = 0 Then Exit Function
Dim U&: U = UB(A)
Dim O$()
ReDim O(U)
Dim J&
For J = 0 To U
    O(J) = RmvPfx(A(J), Pfx)
Next
AyRmvPfx = O
End Function

Function AyRmvSngQRmk(A$()) As String()
If Sz(A) = 0 Then Exit Function
Dim X, O$()
For Each X In Itr(A)
    If Not IsSngQRmk(CStr(X)) Then Push O, X
Next
AyRmvSngQRmk = O
End Function

Function AyRmvSngQuote(A$()) As String()
Dim I
For Each I In Itr(A)
    PushI AyRmvSngQuote, RmvSngQuote(I)
Next
End Function

Function AyRmvT1(A) As String()
Dim I
For Each I In Itr(A)
    PushI AyRmvT1, RmvT1(I)
Next
End Function


Function AyRmvTT(A$()) As String()
Dim I
For Each I In Itr(A)
    PushI AyRmvTT, RmvTT(I)
Next
End Function

Function AyRmv2Dash(Ay) As String()
If Sz(Ay) = 0 Then Exit Function
Dim O$(), I
For Each I In Ay
    Push O, Rmv2Dash(CStr(I))
Next
AyRmv2Dash = O
End Function

