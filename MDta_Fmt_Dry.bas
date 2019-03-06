Attribute VB_Name = "MDta_Fmt_Dry"
Option Explicit
Enum eDryFmt
    eVbarSep
    eSpcSep
End Enum
Sub DmpDry(Dry(), Optional Fmt As eDryFmt)
D FmtDry(Dry, Fmt:=Fmt)
End Sub

Function FmtDry(Dry(), _
Optional MaxColWdt% = 100, Optional BrkColIx% = -1, Optional ShwZer As Boolean, Optional Fmt As eDryFmt) As String()
If Sz(Dry) = 0 Then Exit Function
Dim Dry1(): Dry1 = DryzStrCell(Dry, ShwZer, Fmt, MaxColWdt) ' Convert each cell in Dry-Dry into string
Dim W%(): W = WdtAyzDry(Dry1)
If Fmt = eSpcSep Then
    Dim Dr
    For Each Dr In Dry1
        PushI FmtDry, Jn(Dr)
    Next
    Exit Function
End If

Dim SepDr1$(): SepDr1 = SepDr(W)
Dim Dry2()
    If BrkColIx >= 0 Then
        Dry2 = InsSepDr(Dry1, BrkColIx, SepDr1)
    Else
        Dry2 = Dry1
    End If
Dim SepLin$: SepLin = SepLinzSepDr(SepDr1, Fmt)
Push FmtDry, SepLin
    For Each Dr In Dry2
        PushI FmtDry, FmtDr(Dr, W, Fmt)
    Next
PushI FmtDry, SepLin
End Function

Private Function DryzStrCell(Dry, ShwZer As Boolean, Fmt As eDryFmt, MaxColWdt%) As Variant()
Dim Dr
For Each Dr In Dry
   Push DryzStrCell, DrzStrCell(Dr, ShwZer, Fmt, MaxColWdt)
Next
End Function

Private Function DrzStrCell(Dr, ShwZer As Boolean, Fmt As eDryFmt, MaxWdt%) As String()
Dim I, P$, S$
If Fmt = eVbarSep Then P = " ": S = " "
For Each I In Itr(Dr)
    PushI DrzStrCell, P & LinzVal(I, ShwZer, MaxWdt) & S
Next
End Function
Private Function StrCell$(V, Optional ShwZer As Boolean) ' Convert V into a string in a cell
'SpcSepStr is a string can be displayed in a cell
Select Case True
Case IsNumeric(V)
    If V = 0 Then
        If ShwZer Then
            StrCell = "0"
        End If
    Else
        StrCell = V
    End If
Case IsEmp(V):
Case IsArray(V)
    Dim N&: N = Sz(V)
    If N = 0 Then
        StrCell = "*[0]"
    Else
        StrCell = "*[" & N & "]" & V(0)
    End If
Case IsObject(V): StrCell = TypeName(V)
Case Else:        StrCell = V
End Select
End Function

Private Function InsSepDr(Dry(), BrkColIx%, SepDr$()) As Variant()
Dim Dr, IsBrk As Boolean, LasCell$, Fst As Boolean
LasCell = Dry(0)(BrkColIx)
For Each Dr In Dry
    If Fst Then
        Fst = False
    Else
        IsBrk = LasCell = Dr(BrkColIx)
    End If
    If IsBrk Then
        PushI InsSepDr, SepDr
        LasCell = Dr(BrkColIx)
    End If
    Push InsSepDr, Dr
Next
End Function
Sub BrwDry(A(), Optional MaxColWdt% = 100, Optional BrkColIx% = -1, Optional ShwZer As Boolean, Optional Fmt As eDryFmt)
BrwAy FmtDry(A, MaxColWdt, BrkColIx, ShwZer, Fmt)
End Sub
Private Function WdtAyzDry(A()) As Integer()
Dim J%
For J = 0 To NColzDry(A) - 1
    Push WdtAyzDry, WdtzAy(ColzDry(A, J))
Next
End Function

