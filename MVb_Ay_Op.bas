Attribute VB_Name = "MVb_Ay_Op"
Option Explicit
Const CMod$ = "MVb_Ay__Operation."
Function DashLT1Ay(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushNoDup DashLT1Ay, TakBefOrAll(I, "_")
Next
End Function
Function AyEndTrim(A$()) As String()
If Sz(A) = 0 Then Exit Function
If LasEle(A) <> "" Then AyEndTrim = A: Exit Function
Dim J%
For J = UB(A) To 0 Step -1
    If Trim(A(J)) <> "" Then
        Dim O$()
        O = A
        ReDim Preserve O(J)
        AyEndTrim = O
        Exit Function
    End If
Next
End Function

Function AyIntersect(A, B)
AyIntersect = AyCln(A)
If Sz(A) = 0 Then Exit Function
If Sz(A) = 0 Then Exit Function
Dim V
For Each V In A
    If HasEle(B, V) Then PushI AyIntersect, V
Next
End Function
Function MinAy(A)
Dim O, J&
If Sz(A) = 0 Then Exit Function
O = A(0)
For J = 1 To UB(A)
    If A(J) < O Then O = A(J)
Next
MinAy = O
End Function

Function AyMinus(A, B)
If Sz(B) = 0 Then AyMinus = A: Exit Function
AyMinus = AyCln(A)
If Sz(A) = 0 Then Exit Function
Dim V
For Each V In A
    If Not HasEle(B, V) Then
        PushI AyMinus, V
    End If
Next
End Function

Function AyMinusAp(Ay, ParamArray AyAp())
Dim O: O = Ay
Dim AyAv(): AyAv = AyAp
Dim Ayi
For Each Ayi In AyAv
    If Sz(O) = 0 Then GoTo X
    O = AyMinus(O, Ayi)
Next
X:
AyMinusAp = O
End Function

Function MaxAy(A)
Dim O, I
For Each I In Itr(A)
    If I > O Then O = I
Next
MaxAy = O
End Function

Function MaxAySz%(A)
If Sz(A) = 0 Then Exit Function
Dim O&, I, S&
For Each I In A
    O = Max(O, Sz(I))
Next
MaxAySz = O
End Function


Function Ny(A) As String()
Const CSub$ = CMod & "Ny"
Select Case True
Case IsStr(A): Ny = SySsl(A)
Case IsSy(A): Ny = A
Case Else: ThwPmEr A, CSub, "Should be Str or Sy"
End Select
End Function



Function CvVy(Vy)
Const CSub$ = CMod & "CvVy"
Select Case True
Case IsStr(Vy): CvVy = SySsl(Vy)
Case IsArray(Vy): CvVy = Vy
Case Else: Thw CSub, "VyzDicKK should either be string or array", "Vy-TypeName Vy", TypeName(Vy), Vy
End Select
End Function
Function CvNy(Ny0) As String()
Const CSub$ = CMod & "CvNy"
Select Case True
Case IsMissing(Ny0) Or IsEmpty(Ny0)
Case IsStr(Ny0): CvNy = TermAy(Ny0)
Case IsSy(Ny0): CvNy = Ny0
Case IsArray(Ny0): CvNy = SyzAy(Ny0)
Case Else: Thw CSub, "Given Ny0 must be Missing | Empty | Str | Sy | Ay", "TypeName-Ny0", TypeName(Ny0)
End Select
End Function

Function CvAv(A) As Variant()
CvAv = A
End Function

Function CvSy(A) As String()
Select Case True
Case IsEmpty(A) Or IsMissing(A)
Case IsSy(A): CvSy = A
Case IsArray(A): CvSy = SyzAy(A)
Case Else: CvSy = Sy(CStr(A))
End Select
End Function

Function SyShow(XX$, Sy$()) As String()
Dim O$()
Select Case Sz(Sy)
Case 0
    Push O, XX & "()"
Case 1
    Push O, XX & "(" & Sy(0) & ")"
Case Else
    Push O, XX & "("
    PushAy O, Sy
    Push O, XX & ")"
End Select
SyShow = O
End Function

Private Sub ZZ()
Dim A
Dim B()
Dim C$
Dim D$()
Dim XX
CvNy A
CvSy A
Sy B
SyShow C, D
End Sub

Private Sub Z()
End Sub

