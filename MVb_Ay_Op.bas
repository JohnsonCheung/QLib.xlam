Attribute VB_Name = "MVb_Ay_Op"
Option Explicit
Const CMod$ = "MVb_Ay__Operation."
Function DashLT1Ay(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushNoDup DashLT1Ay, BefOrAll(CStr(I), "_")
Next
End Function

Function AyEndTrim(A$()) As String()
If Si(A) = 0 Then Exit Function
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
If Si(A) = 0 Then Exit Function
If Si(A) = 0 Then Exit Function
Dim V
For Each V In A
    If HasEle(B, V) Then PushI AyIntersect, V
Next
End Function
Function MinAy(A)
Dim O, J&
If Si(A) = 0 Then Exit Function
O = A(0)
For J = 1 To UB(A)
    If A(J) < O Then O = A(J)
Next
MinAy = O
End Function

Function AyMinusAp(Ay, ParamArray Ap())
Dim IAy, O
O = Ay
For Each IAy In Ap
    O = AyMinus(Ay, IAy)
    If Si(O) = 0 Then AyMinusAp = O: Exit Function
Next
AyMinusAp = O
End Function

Function AyMinus(A, B)
If Si(B) = 0 Then AyMinus = A: Exit Function
AyMinus = AyCln(A)
If Si(A) = 0 Then Exit Function
Dim V
For Each V In A
    If Not HasEle(B, V) Then
        PushI AyMinus, V
    End If
Next
End Function

Function MaxAy(A)
Dim O, I
For Each I In Itr(A)
    If I > O Then O = I
Next
MaxAy = O
End Function

Function Ny(S$) As String()
Ny = TermAy(S)
End Function

Function CvVy(Vy)
Const CSub$ = CMod & "CvVy"
Select Case True
Case IsStr(Vy): CvVy = SySsl(CStr(Vy))
Case IsArray(Vy): CvVy = Vy
Case Else: Thw CSub, "VyzDicKK should either be string or array", "Vy-TypeName Vy", TypeName(Vy), Vy
End Select
End Function

Function CvBytAy(A) As Byte()
CvBytAy = A
End Function

Function CvAv(A) As Variant()
CvAv = A
End Function
Function CvObj(A) As Object
Set CvObj = A
End Function
Function CvSy(A) As String()
Select Case True
Case IsSy(A): CvSy = A
Case IsArray(A): CvSy = SyzAy(A)
Case IsEmpty(A) Or IsMissing(A)
Case Else: CvSy = Sy(CStr(A))
End Select
End Function

Function SyShow(XX$, Sy$()) As String()
Dim O$()
Select Case Si(Sy)
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
CvSy A
Sy B
SyShow C, D
End Sub

Private Sub Z()
End Sub

