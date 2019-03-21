VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActMd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const Ns$ = "MdyPj"
Public Md As CodeModule
Private A() As ActLin
Friend Function Init(Md As CodeModule, ActLin() As ActLin) As ActMd
Set Me.Md = Md
A = ActLin
Set Init = Me
End Function

Function ActLinAy() As ActLin()
ActLinAy = A
End Function
Function Hdr$()
Hdr = "Md L# Ins/Dlt Lin"
End Function
Function ToFmt(Optional NoHdr As Boolean) As String()
Dim O$()
If Not NoHdr Then PushI O, Hdr
PushIAy O, ToLy
ToFmt = FmtAyT3(O)
End Function

Function ToLy() As String()
Dim J&
For J = 0 To UB(A)
    With A(J)
    PushI ToLy, MdNm(Md) & " " & .Lno & " " & .ActStr & " " & .Lin
    End With
Next
End Function
