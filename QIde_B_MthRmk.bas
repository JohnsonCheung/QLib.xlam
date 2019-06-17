Attribute VB_Name = "QIde_B_MthRmk"
Option Explicit
Option Compare Text

Sub EnsMthRmk(M As CodeModule, Mthn, NewRmk$)
'Ret : mk sure the rmk of #Mthn will be #NewRmk
Dim MthLno&: MthLno = MthLnozMM(M, Mthn)
If MthLno = 0 Then
    Debug.Print "EnsMthRmk: no such mth[" & Mthn & "]"
    Debug.Print
    Exit Sub
End If
Dim RLno&: RLno = MthRmkLno(M, MthLno)
Dim RmkL$: RmkL = MthRmkLzRmkLno(M, RLno)

Select Case True
Case RmkL = "" And NewRmk <> ""
    RLno = NxtLnozML(M, MthLno)
    GoSub Ins
Case RmkL = ""
    Debug.Print "EnsMthRmk: mth[" & Mthn & "] has no rmk and no new rmk"
Case RmkL = NewRmk
    Debug.Print "EnsMthRmk: Same": Exit Sub
Case Else
    DltLines M, RLno, RmkL
    If NewRmk <> "" Then GoSub Ins
End Select
Exit Sub
Ins:
    Debug.Print "EnsMthRmk: Inserted.  Mthn[" & Mthn & "]"
    Debug.Print BoxLines(NewRmk)
    Debug.Print
    M.InsertLines RLno, NewRmk
    Return

End Sub

Sub EnsMthRmkzS1S2(M As CodeModule, NewRmk As S1S2s)
Dim Ay() As S1S2: Ay = NewRmk.Ay
Dim J%
For J = 0 To NewRmk.N - 1
    EnsMthRmk M, Ay(J).S1, Ay(J).S2
Next
End Sub

Private Function MthRmkIx&(Src$(), MthIx)
If IsSngLinMth(Src(MthIx)) Then Exit Function
Dim ELin$: ELin = EndLin(Src, MthIx)
Dim Ix&: For Ix = MthIx + 1 To UB(Src)
    Dim L$: L = Src(Ix)
    If IsVbRmk(L) Then MthRmkIx = Ix: Exit Function
    If L = ELin Then Exit Function
Next
ThwImpossible CSub
End Function

Private Function MthRmkzMthIx$(Src$(), MthIx)
Dim R&: R = MthRmkIx(Src, MthIx)
MthRmkzMthIx = MthRmkLzRmkIx(Src, R)
End Function

Private Function MthRmkL$(M As CodeModule, MthLno&)
MthRmkL = MthRmkLzRmkLno(M, MthRmkLno(M, MthLno))
End Function

Private Function MthRmkLzRmkIx$(Src$(), RmkIx&)
Dim RBlk$(): RBlk = RmkBlkzS(Src, RmkIx)
Dim Adj$(): Adj = RAdj(RBlk)
MthRmkLzRmkIx = JnCrLf(Adj)
End Function

Private Function MthRmkLzRmkLno$(M As CodeModule, RLno&)
Dim RBlk$(): RBlk = RmkBlkzM(M, RLno)
Dim Adj$(): Adj = RAdj(RBlk)
MthRmkLzRmkLno = JnCrLf(Adj)
End Function

Private Function MthRmkLno&(M As CodeModule, MthLno&)
Dim ELin$: ELin = EndLinzM(M, MthLno)
Dim Lno&: For Lno = MthLno + 1 To M.CountOfLines
    Dim L$: L = M.Lines(Lno, 1)
    If IsVbRmk(L) Then MthRmkLno = Lno: Exit Function
    If L = ELin Then Exit Function
Next
ThwImpossible CSub
End Function

Sub Z_MthRmkzM()
Dim M As CodeModule: Set M = Md("QIde_B_MthOp__AlignMthDimzML")
BrwS1S2s MthRmkzM(M)
End Sub

Function MthRmkzNy(M As CodeModule, MthNy$()) As S1S2s
Dim S$(): S = Src(M)
Dim MthIxy&(): MthIxy = MthIxyzSNy(S, MthNy)
MthRmkzNy = MthRmkzMthIxy(Src(M), MthIxy)
End Function

Private Function MthRmkzMthIxy(Src$(), MthIxy&()) As S1S2s
Dim Ix, O As S1S2s: For Each Ix In Itr(MthIxy)
    Dim R$: R = MthRmkzMthIx(Src, Ix)
    If R <> "" Then
        Dim N$: N = Mthn(Src(Ix))
        PushS1S2 O, S1S2(N, R)
    End If
Next
MthRmkzMthIxy = O
End Function

Function MthRmkP() As S1S2s
MthRmkP = MthRmkzP(CPj)
End Function

Sub Z_MthRmkP()
BrwS1S2s MthRmkP
End Sub

Function MthRmkzP(P As VBProject) As S1S2s
Dim C As VBComponent
For Each C In P.VBComponents
    Dim A As S1S2s: A = MthRmkzM(C.CodeModule)
    Dim B As S1S2s: B = AddS1Pfx(A, C.Name & ".")
    PushS1S2s MthRmkzP, B
Next
End Function

Function MthRmkzM(M As CodeModule) As S1S2s
Dim S$(): S = Src(M)
Dim Ixy&(): Ixy = MthIxy(S)
MthRmkzM = MthRmkzMthIxy(S, Ixy)
End Function

Sub Z_EnsMthRmk()
'GoSub Z1
Dim M As CodeModule
GoSub Z1
Exit Sub
Z1:
    'GoSub Crt: Exit Sub
    Set M = Md("TmpMod123")
    EnsMthRmk M, "AAXX", "'skldfjsdlkfj lksdj flksdj fkj @@"
    Return
Z2:
    Set M = Md("TmpMod20190605_231101")
    EnsMthRmk M, "AAXX", RplVBar("'sldkfjsd|'slkdfj|slkdfj|'sldkfjsdf|'sdf")
    Return
Z3:
    Set M = Md("TmpMod20190605_231101")
    EnsMthRmk M, "AAXX", RplVBar("'a|'bb|'cfsdfdsc")
    Return
Crt:
    EnsMod CPj, "TmpMod123"
    Set M = Md("TmpMod123")
    ClrMd M
    M.AddFromString "Sub AAXX()" & vbCrLf & "End Sub"
    Return
End Sub

Private Function RAdj(RBlk$()) As String()
Dim O$(), L: For Each L In Itr(RBlk)
    PushI O, L
    If Right(L, 2) = "@@" Then RAdj = O: Exit Function
Next
End Function

Private Function RmkBlkzM(M As CodeModule, RLno&) As String()
If RLno = 0 Then Exit Function
Dim J&, L$, O$()
For J = RLno To M.CountOfLines
    L = M.Lines(J, 1)
    If Not IsVbRmk(L) Then Exit For
    PushI O, L
Next
RmkBlkzM = O
End Function

Private Function RmkBlkzS(Src$(), RmkIx&) As String()
If RmkIx = 0 Then Exit Function
Dim J&, L$, O$()
For J = RmkIx To UB(Src)
    L = Src(J)
    If Not IsVbRmk(L) Then Exit For
    PushI O, L
Next
RmkBlkzS = O
End Function

Private Sub Z()
QIde_B_MthRmk:
End Sub
