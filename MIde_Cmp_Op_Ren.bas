Attribute VB_Name = "MIde_Cmp_Op_Ren"
Option Explicit

Sub RmvModPfx(Pj As VBProject, Pfx$)
Dim C As VBComponent
For Each C In Pj.VBComponents
    If HasPfx(C.Name, Pfx) Then
        RenCmp C, RmvPfx(C.Name, Pfx)
    End If
Next
End Sub

Sub RplModPfx(FmPfx$, ToPfx$)
RplModPfxzPj CurPj, FmPfx, ToPfx
End Sub

Sub RenCmp(A As VBComponent, NewNm$)
If HasCmp(NewNm) Then
    InfLin CSub, "New cmp exists", "OldCmp NewCmp", A.Name, NewNm
Else
    A.Name = NewNm
End If
End Sub

Sub RplModPfxzPj(Pj As VBProject, FmPfx$, ToPfx$)
Dim C As VBComponent, N$
For Each C In Pj.VBComponents
    If C.Type = vbext_ct_StdModule Then
        If HasPfx(C.Name, FmPfx) Then
            RenCmp C, RplPfx(C.Name, FmPfx, ToPfx)
        End If
    End If
Next
End Sub

Sub AddCmpSfxPj(Sfx)
AddCmpSfx CurPj, Sfx
End Sub
Sub AddCmpSfx(A As VBProject, Sfx)
If A.Protection = vbext_pp_locked Then Exit Sub
Dim C As VBComponent
For Each C In A.VBComponents
    RenCmp C, C.Name & Sfx
Next
End Sub

Function SetCmpNm(A As VBComponent, Nm, Optional Fun$ = "SetCmpNm") As VBComponent
Dim Pj As VBProject
Set Pj = PjzCmp(A)
If HasCmpzPj(Pj, Nm) Then
    Thw Fun, "Cmp already Has", "Cmp Has-in-Pj", Nm, Pj.Name
End If
If Pj.Name = Nm Then
    Thw Fun, "CmpNm same as PjNm", "CmpNm", Nm
End If
A.Name = Nm
Set SetCmpNm = A
End Function


