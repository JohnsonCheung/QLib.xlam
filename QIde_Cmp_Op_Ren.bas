Attribute VB_Name = "QIde_Cmp_Op_Ren"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Cmp_Op_Ren."
Private Const Asm$ = "QIde"

Sub RmvModPfx(Pj As VBProject, Pfx$)
Dim C As VBComponent
For Each C In Pj.VBComponents
    If HasPfx(C.Name, Pfx) Then
        RenCmp C, RmvPfx(C.Name, Pfx)
    End If
Next
End Sub

Sub RplModPfx(FmPfx$, ToPfx$)
RplModPfxzP CPj, FmPfx, ToPfx
End Sub

Sub RenCmp(A As VBComponent, NewNm$)
If HasCmpzN(NewNm) Then
    InfLin CSub, "New cmp exists", "OldCmp NewCmp", A.Name, NewNm
Else
    A.Name = NewNm
End If
End Sub

Sub RplModPfxzP(Pj As VBProject, FmPfx$, ToPfx$)
Dim C As VBComponent, N$
For Each C In Pj.VBComponents
    If C.Type = vbext_ct_StdModule Then
        If HasPfx(C.Name, FmPfx) Then
            RenCmp C, RplPfx(C.Name, FmPfx, ToPfx)
        End If
    End If
Next
End Sub

Sub AddCmpSfxP(Sfx)
AddCmpSfx CPj, Sfx
End Sub
Sub AddCmpSfx(P As VBProject, Sfx)
If P.Protection = vbext_pp_locked Then Exit Sub
Dim C As VBComponent
For Each C In P.VBComponents
    RenCmp C, C.Name & Sfx
Next
End Sub

Function SetCmpNm(A As VBComponent, Nm, Optional Fun$ = "SetCmpNm") As VBComponent
Dim Pj As VBProject
Set Pj = PjzC(A)
If HasCmpzP(Pj, Nm) Then
    Thw Fun, "Cmp already Has", "Cmp Has-in-Pj", Nm, Pj.Name
End If
If Pj.Name = Nm Then
    Thw Fun, "Cmpn same as Pjn", "Cmpn", Nm
End If
A.Name = Nm
Set SetCmpNm = A
End Function


