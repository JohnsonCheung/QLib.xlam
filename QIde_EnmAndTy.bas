Attribute VB_Name = "QIde_EnmAndTy"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_EnmAndTy."
Private Const Asm$ = "QIde"

Function MdPosStr$(A As MdPos)
Dim B$
With A
    'With .LinPos.Pos
        'If .Cno1 > 0 Then B = " " & .Cno1 & " " & .Cno2
    'End With
    'MdPosStr = "MdPos " & Mdn(A.Md) & A.LinPos.Lno & B
End With
End Function

Function MdPoszMLCC(Md As CodeModule, L, Cno1, Cno2) As MdPos
'MdPoszMLCC = MdPos(Md, LinPoszLCC(L, Cno1, Cno2))
End Function

Function MdPoszMLP(Md As CodeModule, Lno, P As Pos) As MdPos
'MdPoszMLP = MdPos(Md, LinPos(Lno, P))
End Function

Function MdPos(Md As CodeModule, RRCC As RRCC) As MdPos
Set MdPos.Md = Md
MdPos.RRCC = RRCC
End Function

