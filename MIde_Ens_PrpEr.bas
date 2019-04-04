Attribute VB_Name = "MIde_Ens_PrpEr"
Option Explicit
Const CMod$ = "MIde_Ens_PrpEr."

Private Sub EnsLinzExit(A As CodeModule, PrpLno&)
Const CSub$ = CMod & "EnsLinzExit"
Dim L&
L = LnozInsExit(A, PrpLno)
If L = 0 Then Exit Sub
A.InsertLines L, "Exit Property"
Inf CSub, "Exit Property is inserted", "Md PrpLno At", MdNm(A), PrpLno, L
End Sub

Private Sub EnsLinzLblX(A As CodeModule, PrpLno&)
Const CSub$ = CMod & "EnsLinzLblX"
Dim E$, L%, ActLblXLin$, EndPrpLno&
E = LinzLblX(A, PrpLno)
L = LnozLblX(A, PrpLno)
If L <> 0 Then
    ActLblXLin = A.Lines(L, 1)
End If
If E <> ActLblXLin Then
    If L = 0 Then
        EndPrpLno = LnozEndPrp(A, PrpLno)
        If EndPrpLno = 0 Then Stop
        A.InsertLines EndPrpLno, E
        Inf CSub, "Inserted [at] with [line]", EndPrpLno, E
    Else
        A.ReplaceLine L, E
        Inf CSub, "Replaced [at] with [line]", L, E
    End If
End If
End Sub

Private Sub EnsPrpOnErzLno(A As CodeModule, PrpLno&)
If HasSubStr(A.Lines(PrpLno, 1), "End Property") Then
    Exit Sub
End If
EnsLinzLblX A, PrpLno
EnsLinzExit A, PrpLno
EnsLinzOnEr A, PrpLno
End Sub

Private Sub EnsLinzOnEr(A As CodeModule, PrpLno&)
Const CSub$ = CMod & "EnsLinzOnEr"
Dim L&
L = IxOfOnEr(A, PrpLno)
If L <> 0 Then Exit Sub
A.InsertLines PrpLno + 1, "On Error Goto X"
'If Trc Then Msg CSub, "Exit Property is inserted [at]", L
End Sub

Private Function LnozExit&(A As CodeModule, PrpLno)
If HasSfx(A.Lines(PrpLno, 1), "End Property") Then Exit Function
Dim J%, L$
For J = PrpLno + 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If HasPfx(L, "Exit Property") Then LnozExit = J: Exit Function
    If HasPfx(L, "End Property") Then Exit Function
Next
Stop
End Function

Private Function LnozInsExit&(A As CodeModule, PrpLno)
If LnozExit(A, PrpLno) <> 0 Then Exit Function
Dim L%
L = LnozLblX(A, PrpLno)
If L = 0 Then Stop
LnozInsExit = L
End Function

Private Function LinzLblX$(A As CodeModule, PrpLno)
Dim Nm$, Lin$
Lin = A.Lines(PrpLno, 1)
Nm = PrpNm(Lin)
If Nm = "" Then Stop
LinzLblX = FmtQQ("X: Debug.Print ""?.?.PrpEr...[""; Err.Description; ""]""", MdNm(A), Nm)
End Function

Private Function LnozLblX&(A As CodeModule, PrpLno)
Dim J&, L$
For J = PrpLno + 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If HasPfx(L, "X: Debug.Print") Then LnozLblX = J: Exit Function
    If HasPfx(L, "End Property") Then Exit Function
Next
Stop
End Function

Private Function IxOfOnEr&(A As CodeModule, PrpLno)
Dim J%, L$
For J = PrpLno + 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If HasPfx(L, "On Error Goto X") Then IxOfOnEr = J: Exit Function
    If HasPfx(L, "End Property") Then Exit Function
Next
Stop '
End Function

Private Function LnozEndPrp&(A As CodeModule, PrpLno)
If HasSfx(A.Lines(PrpLno, 1), "End Property") Then LnozEndPrp = PrpLno: Exit Function
Dim J%
For J = PrpLno + 1 To A.CountOfLines
    If HasPfx(A.Lines(J, 1), "End Property") Then LnozEndPrp = J: Exit Function
Next
Stop
End Function

Private Sub EnsPrpOnErzMd(A As CodeModule)
Dim J%, L&()
L = PurePrpLnoAy(A)
ThwAyNotSrt L, CSub
For J = UB(L) To 0 Step -1
    EnsPrpOnErzLno A, L(J)
Next
End Sub

Private Sub RmvPrpOnErzLno(A As CodeModule, PrpLno&)
RmvMdLno A, LnozExit(A, PrpLno)
RmvMdLno A, IxOfOnEr(A, PrpLno)
RmvMdLno A, LnozLblX(A, PrpLno)
End Sub

Private Sub RmvPrpOnErzMd(A As CodeModule)
If A.Parent.Type <> vbext_ct_ClassModule Then Exit Sub
Dim J%, L&()
L = PurePrpLnoAy(A)
If Not IsSrtAy(L) Then Stop
For J = UB(L) To 0 Step -1
    RmvPrpOnErzLno A, L(J)
Next
End Sub

Sub RmvPrpOnErOfMd()
RmvPrpOnErzMd CurMd
End Sub

Sub EnsPrpOnErOfMd()
EnsPrpOnErzMd CurMd
End Sub
Private Sub Z_EnsPrpOnErzMd()
'EnsPrpOnErzMd ZZMd
End Sub


