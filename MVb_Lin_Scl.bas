Attribute VB_Name = "MVb_Lin_Scl"
Option Explicit
Const CMod$ = "MVb_Lin_Scl."

Sub AsgSclNN(Scl$, NN$, ParamArray OAp())
Const CSub$ = CMod & "AsgSclNN"
Dim Av(): Av = OAp
Dim V, Ny$(), I, J%
Ny = NyzNN(NN)
If Sz(Ny) <> Sz(Av) Then Stop
For Each I In Itr(AyeEmpEle(AyTrim(SplitSemi(Scl))))
    V = SclItm_V(CStr(I), Ny)
    Select Case True
    Case IsByt(V) And (V = 1 Or V = 2)
    Case IsBool(V) Or IsStr(V): OAp(J) = V
    Case Else: Thw CSub, "Program error in SclItm_V.  It should return one of (Byt1,Byt2,Bool,Str)", "[But now it returns]", TypeName(V)
    End Select
    J = J + 1
Next
End Sub

Function ChkSclNN(A$, Ny0) As String()
Const CSub$ = CMod & "ChkSclNN"
Dim V, Ny$(), I, Er1$(), Er2$()
Ny = NyzNN(Ny0)
For Each I In Itr(AyeEmpEle(AyTrim(SplitSemi(A))))
    V = SclItm_V(CStr(I), Ny)
    Select Case True
    Case IsByt(V) And V = 1: Push Er1, I
    Case IsByt(V) And V = 2: Push Er2, I
    Case IsBool(V) Or IsStr(V)
    Case Else: Thw CSub, "Program error in SclItm.  It should return (Byt1,Byt2,Bool,Str), but now it returns [Ty]", TypeName(V)
    End Select
Next
Dim O$()
    If Sz(Er1) > 0 Then
        O = LyzMsgNap("There are [invalid-SclNy] in given [scl] under these [valid-SclNy].", "Er Ny", JnSpc(Er1), A, JnSpc(Ny))
    End If
    If Sz(Er2) > 0 Then
        PushAy O, LyzMsgNap("[Itm] of [Scl] has [valid-SclNy], but it is not one of SclNy nor it has '='", "Er Scl Valid-SclNy", Er2, A, Ny)
    End If
ChkSclNN = O
End Function

Function SclItm_V(A$, Ny$())
'Return Byt1 if Pfx of A not in Ny
'Return True If A = One Of Ny
'Return Byt2 if Pfx of A is in Ny, but not Eq one Ny and Don't have =
If HasEle(Ny, A) Then SclItm_V = True: Exit Function
If Not HasStrPfxAy(A, Ny) Then SclItm_V = CByte(1): Exit Function
If Not HasSubStr(A, "=") Then SclItm_V = CByte(2): Exit Function
SclItm_V = Trim(TakAft(A, "="))
End Function

Function ShfScl$(OStr$)
AsgBrk1 OStr, ";", ShfScl, OStr
End Function

Private Sub ZZ()
Dim A$
Dim B As Variant
Dim C()
Dim D$()
End Sub

Private Sub Z()
End Sub
