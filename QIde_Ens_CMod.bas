Attribute VB_Name = "QIde_Ens_CMod"
Option Explicit
Private Const CMod$ = "BEnsCMod."
Private Const Asm$ = "QIde"
Private Const Ns$ = "QIde.Qualify"
Private Sub Z_EnsCModzMd()
Dim Md As CodeModule
GoSub T0
Exit Sub
T0:
    Set Md = CurMd
    GoTo Tst
Tst:
    EnsCModzMd Md
    Return
End Sub

Sub EnsCModM()
EnsCModzMd CurMd
End Sub

Sub EnsCModP()
EnsCModzPj CurPj
End Sub

Sub EnsCModzPj(A As VBProject)
Dim C As VBComponent
For Each C In A.VBComponents
    EnsCModzMd C.CodeModule
Next
End Sub
Sub EnsCModzMd(A As CodeModule)
With SomMdygLinPmzSetgCModConst(A)
    If .Som Then
        Debug.Print MdNm(A); "<============= Mdy"
        MdyMdzLin A, .Itm
    End If
End With
End Sub

Function ConstLinzCMod$(A As CodeModule)
Dim N$: N = MdNm(A)
If N = "" Then Exit Function
ConstLinzCMod = FmtQQ("Private Const CMod$ = ""?.""", N)
End Function

Function LnozCModConst(A As CodeModule)
LnozCModConst = LnozConst(A, "CMod$")
End Function

Function SomMdygLinPmzSetgCModConst(A As CodeModule) As SomMdygLinPm
Dim NewLin$: NewLin = ConstLinzCMod(A)
Dim O As SomMdygLinPm
Dim Lno&: Lno = LnozCModConst(A)
Dim OldLin$: OldLin = ContLinzMd(A, Lno)
Select Case True
Case Lno = 0: O = SomInsgLinPm(LnozAftOptzAndImpl(A), NewLin)
Case Lno > 0 And OldLin = "": Thw CSub, "Lno>0, OldLin must have value", "Md Lno", MdNm(A), Lno
Case Lno > 0 And OldLin = NewLin:
Case Lno > 0 And OldLin <> NewLin: O = SomRplgLinPm(Lno, OldLin, NewLin)
Case Else: ThwImpossible CSub
End Select
SomMdygLinPmzSetgCModConst = O
End Function

Private Sub Z_SomMdygLinPmzSetgCModConst()
Dim Md As CodeModule, Act As SomMdygLinPm, Ept As SomMdygLinPm
GoSub T0
Exit Sub
T0:
    Set Md = CurMd
    Ept = SomInsgLinPm(2, "Private Const CMod$ = ""BEnsCMod.""")
    GoTo Tst
Tst:
    Act = SomMdygLinPmzSetgCModConst(Md)
    If Not IsEqSomMdygLinPm(Act, Ept) Then Stop
    Return
End Sub

