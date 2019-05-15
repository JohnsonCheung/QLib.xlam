Attribute VB_Name = "QVb_Dic_NewDic"
Option Explicit
Private Const CMod$ = "MVb_Dic_New."
Private Const Asm$ = "QVb"
Function DiczFt(Ft) As Dictionary
Set DiczFt = Dic(LyzFt(Ft))
End Function
Function NewSyDic(TermLiny$()) As Dictionary
Dim L$, I, T$, Ssl$
Dim O As New Dictionary
For Each I In Itr(TermLiny)
    L = I
    AsgTRst L, T, Ssl
    If O.Exists(T) Then
        O(T) = AddAy(O(T), SyzSS(Ssl))
    Else
        O.Add T, SyzSS(Ssl)
    End If
Next
Set NewSyDic = O
End Function

Sub DicSetKv(O As Dictionary, K, V)
If O.Exists(K) Then
    Asg V, O(K)
Else
    O.Add K, V
End If
End Sub

Sub AddDiczNonBlankStr(ODic As Dictionary, K, S$)
If S = "" Then Exit Sub
ODic.Add K, S
End Sub

Function DiczLines(Lines$, Optional JnSep$ = vbCrLf) As Dictionary
Set DiczLines = Dic(SplitCrLf(Lines), JnSep)
End Function

Sub ApdLinzToLinesDic(OLinesDic As Dictionary, K, Lin, Sep$)
If OLinesDic.Exists(K) Then
    OLinesDic(K) = OLinesDic(K) & Sep & Lin
Else
    OLinesDic.Add K, Lin
End If
End Sub

Function LyzLinesDicItems(LineszDic As Dictionary) As String()
Dim Lines$, I
For Each I In LineszDic.Items
    Lines = I
    PushIAy LyzLinesDicItems, SplitCrLf(Lines)
Next
End Function

Function DiczVkkLy(VkkLy$()) As Dictionary
Set DiczVkkLy = New Dictionary
Dim I, V$, Vkk$, K
For Each I In Itr(VkkLy)
    Vkk = I
    V = T1(Vkk)
    For Each K In SyzSS(RmvT1(Vkk))
        DiczVkkLy.Add K, V
    Next
Next
End Function
Function LyzDic(A As Dictionary) As String()
Dim K
For Each K In A.Keys
    PushI LyzDic, K & " " & A(K)
Next
End Function
Function JnStrDic$(StrDic As Dictionary, Optional Sep$)
JnStrDic = Join(SyzItr(StrDic.Items), Sep)
End Function
Function Dic(Ly$(), Optional JnSep$ = vbCrLf) As Dictionary
Dim O As New Dictionary
Dim I, L$, T$, Rst$
For Each I In Itr(Ly)
    L = I
    AsgTRst L, T, Rst
    If T <> "" Then
        If O.Exists(T) Then
            O(T) = O(T) & JnSep & Rst
        Else
            O.Add T, Rst
        End If
    End If
Next
Set Dic = O
End Function
Function DiczKyVy(Ky, Vy) As Dictionary
ThwIf_DifSi Ky, Vy, CSub
Dim J&
Set DiczKyVy = New Dictionary
For J = 0 To UB(Ky)
    DiczKyVy.Add Ky(J), Vy(J)
Next
End Function

Function DiczVbl(Vbl$, Optional JnSep$ = vbCrLf) As Dictionary
Set DiczVbl = Dic(SplitVBar(Vbl), JnSep)
End Function

Function DiczAyab(A, B) As Dictionary
ThwIf_DifSi A, B, CSub
Dim N1&, N2&
N1 = Si(A)
N2 = Si(B)
If N1 <> N2 Then Stop
Set DiczAyab = New Dictionary
Dim J&, X
For Each X In Itr(A)
    DiczAyab.Add X, B(J)
    J = J + 1
Next
End Function

