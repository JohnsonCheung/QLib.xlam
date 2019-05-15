Attribute VB_Name = "QIde_Mdy_Mdy"
Enum EmMdyg
    EiNop
    EiIns
    EiDlt
End Enum
Type Insg
    Lno As Long
    Lin As String
End Type
Type Dltg
    Lno As Long
    Lin As String
End Type
Type Mdyg
    Act As EmMdyg
    Ins As Insg
    Dlt As Dltg
End Type
Type Mdygs: N As Long: Ay() As Mdyg: End Type
Type RplgMd: Md As CodeModule: NewLines As String: End Type
Type SomRplgMd: Som As Boolean: RplgMd As RplgMd: End Type
Type RplgMds: N As Long: Ay() As RplgMd: End Type
Type RplgPj: Pj As VBProject: RplgMds As RplgMds: End Type

Function AddMdygs(A As Mdygs, B As Mdygs) As Mdygs
AddMdygs = A
Dim J&
For J = 0 To B.N - 1
    PushMdyg AddMdygs, B.Ay(J)
Next
End Function

Sub PushMdygs(O As Mdygs, M As Mdygs)
Dim J&
For J = 0 To M.N - 1
    PushMdyg O, M.Ay(J)
Next
End Sub

Sub PushMdyg(O As Mdygs, M As Mdyg)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub

Function SngMdyg(A As Mdyg) As Mdygs
PushMdyg SngMdyg, A
End Function

Sub MdyMds(A As RplgMds)
Dim J%
For J = 0 To A.N - 1
    MdyMd A.Ay(J)
Next
End Sub

Sub PushRplgMd(O As RplgMds, M As RplgMd)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub

Function FmtRplgMds(A As RplgMds) As String()
Dim J%
For J = 0 To A.N - 1
    PushI FmtRplgMds, FmtRplgMd(A.Ay(J))
Next
End Function

Sub VcRplgMds(A As RplgMds)
Vc FmtRplgMds(A)
End Sub

Function FmtRplgMd$(A As RplgMd)
PushI FmtRplgMd, "Md=" & Mdn(A.Md) & vbCrLf & A.NewLines
End Function
Function FmtMdygs(A As Mdygs) As String()
For J = 0 To A.N - 1
    PushI FmtMdygs, FmtMdyg(A.Ay(J))
Next
Stop
End Function
Function FmtDltg$(A As Dltg)
With A
    FmtDltg = "Dlt " & .Lno & " " & .Lin
End With
End Function
Function FmtInsg$(A As Insg)
With A
    FmtInsg = "Ins " & .Lno & " " & .Lin
End With
End Function
Function FmtMdyg$(A As Mdyg)
Dim O$
With A
Select Case True
Case .Act = EiDlt: O = FmtDltg(.Dlt)
Case .Act = EiIns: O = FmtInsg(.Ins)
Case .Act = EiNop: O = "Nop"
Case Else: ThwImpossible CSub
End Select
End With
FmtMdyg = O
Debug.Print O
Stop
End Function

Sub BrwRplgMds(A As RplgMds)
BrwLinesAy FmtRplgMds(A)
End Sub

Sub BrwRplgMd(A As RplgMd)
B FmtRplgMd(A)
End Sub

Function RplgMd(M As CodeModule, NewLines$) As RplgMd
Set RplgMd.Md = M
RplgMd.NewLines = NewLines
End Function

