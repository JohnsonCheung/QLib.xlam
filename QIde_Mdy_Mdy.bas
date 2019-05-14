Attribute VB_Name = "QIde_Mdy_Mdy"
Enum EmMdyg
    EiNop
    EiIns
    EiDlt
End Enum
Type InsgLin
    Lno As Long
    Lin As String
End Type
Type DltgLin
    Lno As Long
    OldLin As String
End Type
Type Mdyg
    Act As EmMdyg
    Ins As InsgLin
    Dlt As DltgLin
End Type
Type Mdygs: N As Integer: Ay() As Mdyg: End Type
Type MdygMd: Md As CodeModule: MdgyLins As Mdygs: End Type
Type MdygMds: N As Integer: Ay() As MdygMd: End Type
Type MdygPj: Pj As VBProject: MdygMds As MdygMds: End Type

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

Sub MdyMds(A As MdygMds)
Dim J%
For J = 0 To A.N - 1
    MdyMd A.Ay(J)
Next
End Sub

Sub PushMdygMd(O As MdygMds, M As MdygMd)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub

Function LyzMdygMds(A As MdygMds) As String()
Dim J%
For J = 0 To A.N - 1
    PushIAy LyzMdygMds, LyzMdygMd(A.Ay(J))
Next
End Function
Sub VcMdygMds(A As MdygMds)
Vc FmtMdygMds(A)
End Sub
Function FmtMdygMd(A As MdygMd) As String()
PushI FmtMdygMd, "Md=" & Mdn(A.Md)
PushIAy FmtMdygMd, FmtMdygs(A.MdgyLins)
End Function
Function FmtMdygs(A As Mdygs) As String()
For J = 0 To A.N - 1
    PushI FmtMdygs, FmtMdyg(A.Ay(J))
Next
Stop
End Function
Function FmtDltgLin$(A As DltgLin)
With A
    FmtDltgLin = "Dlt " & .Lno & " " & .OldLines
End With
End Function
Function FmtInsgLin$(A As InsgLin)
With A
    FmtInsgLin = "Ins " & .Lno & " " & .Lines
End With
End Function
Function FmtMdyg$(A As Mdyg)
Dim O$
With A
Select Case True
Case .Act = EiDlt: O = FmtDltgLin(.Dlt)
Case .Act = EiIns: O = FmtInsgLin(.Ins)
Case .Act = EiNop: O = "Nop"
Case .Act = EiRpl: O = FmtRplgLin(.Rpl)
Case Else: ThwImpossible CSub
End Select
End With
FmtMdyg = O
Debug.Print O
Stop
End Function

Function FmtMdygMds(A As MdygMds) As String()
Dim J&
For J = 0 To A.N - 1
    PushIAy FmtMdygMds, FmtMdygMd(A.Ay(J))
Next
End Function
Sub BrwMdygMds(A As MdygMds)
B FmtMdygMds(A)
End Sub

Sub BrwMdygMd(A As MdygMd)
B FmtMdygMd(A)
End Sub

Function MdygMd(A As CodeModule, B As Mdygs) As MdygMd
Set MdygMd.Md = A
MdygMd.MdgyLins = B
End Function

