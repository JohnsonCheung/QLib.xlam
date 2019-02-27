VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A_DtAy() As DT
Public DsNm$

Property Get DtAy() As DT()
DtAy = A_DtAy
End Property

Property Get Init(A() As DT, Optional DsNm$ = "Ds") As Ds
A_DtAy = A
Me.DsNm = DsNm
Set Init = Me
End Property

Sub Brw(Optional MaxColWdt% = 100, Optional DtBrkColDicVbl$, Optional NoIxCol As Boolean)
BrwAy Fmt(MaxColWdt, DtBrkColDicVbl, NoIxCol)
End Sub

Sub Dmp()
DmpAy Fmt
End Sub
Property Get UDt%()
UDt = UB(A_DtAy)
End Property
Function DT(Ix%) As DT
Set DT = A_DtAy(Ix)
End Function
Function Fmt(Optional MaxColWdt% = 100, Optional DtBrkColDicVbl$, Optional NoIxCol As Boolean) As String()
Push Fmt, "*Ds " & DsNm & " " & String(10, "=")
Dim Dic As Dictionary
    Set Dic = DiczVbl(DtBrkColDicVbl)
Dim J%, D As DT, BrkColNm$
For J = 0 To UDt
    Set D = DT(J)
    If Dic.Exists(D.DtNm) Then BrkColNm = Dic(D.DtNm) Else BrkColNm = ""
    PushAy Fmt, FmtDt(D, MaxColWdt, BrkColNm, NoIxCol)
Next
End Function
