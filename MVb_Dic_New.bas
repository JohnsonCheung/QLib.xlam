Attribute VB_Name = "MVb_Dic_New"
Option Explicit
Function DiczFt(Ft) As Dictionary
Set DiczFt = Dic(FtLy(Ft))
End Function
Function NewSyDic(TermLinAy$()) As Dictionary
Dim L, T$, Ssl$
Dim O As New Dictionary
For Each L In Itr(TermLinAy)
    AsgTRst L, T, Ssl
    If O.Exists(T) Then
        O(T) = AyAdd(O(T), SySsl(Ssl))
    Else
        O.Add T, SySsl(Ssl)
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

Function FmtDic1(A As Dictionary) As String()
If A.Count = 0 Then Exit Function
Dim Key: Key = A.Keys
Dim O$(): O = FmtAyAlign(Key)
Dim J&
For J = 0 To UB(Key)
   O(J) = O(J) & " " & A(Key(J))
Next
FmtDic1 = O
End Function

Sub AddDiczNonBlankStr(ODic As Dictionary, K, S$)
If S = "" Then Exit Sub
ODic.Add K, S
End Sub

Function DiczLines(Lines, Optional JnSep$ = vbCrLf) As Dictionary
Set DiczLines = Dic(SplitCrLf(Lines), JnSep)
End Function

Function Dic(Ly$(), Optional JnSep$ = vbCrLf) As Dictionary
Dim O As New Dictionary
Dim I, T$, Rst$
For Each I In Itr(Ly)
    AsgTRst I, T, Rst
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
ThwDifSz Ky, Vy, CSub
Dim J&
Set DiczKyVy = New Dictionary
For J = 0 To UB(Ky)
    DiczKyVy.Add Ky(J), Vy(J)
Next
End Function

Function DiczVbl(Vbl, Optional JnSep$ = vbCrLf) As Dictionary
Set DiczVbl = Dic(SplitVBar(Vbl), JnSep)
End Function

