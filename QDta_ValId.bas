Attribute VB_Name = "QDta_ValId"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_ValId."
Private Const Asm$ = "QDta"
Function AddColzValIdqCnt(D As Drs, Coln$, Optional ColnPfx$) As Drs
'Fm D       : ..@Coln..  ! must have a Str-Col-@Coln
'Fm Coln    : Str-Col-Nm !
'Fm ColnPfx :            ! to fnd: %F1 and %F2, where %F1 = %P%Id & %F2 = %P%Cnt where %P = @ColnPfx & @Coln
'Ret        : ..%F1..%F2 ! Add 2 col: %F1 & %F2 to end of @D.
'                        ! %F1 is Id-Col running fm 1 for each dist val of ^@Coln
'                        ! %F2 is Cnt-Col is the cnt of occurrance such id.  Rec of sam @Coln-Val will have sam *Cnt
Dim P$:       P = ColnPfx & Coln                    ' Pfx
Dim F1$:     F1 = P & "Id"                          ' Fld-1
Dim F2$:     F2 = P & "Cnt"                         ' Fld-2
Dim Fny$(): Fny = AddSy(D.Fny, Sy(F1, F2))          ' New-Fny
Dim Ix&:     Ix = IxzAy(Fny, Coln, ThwEr:=EiThwEr) ' Ix-of-Coln

AddColzValIdqCnt = Drs(Fny, AddColzValIdqCnt_ToDy(D.Dy, Ix))
End Function

Private Function AddColzValIdqCnt_ToDy(Dy(), ValCix&) As Variant()
Dim Col():                       Col = ColzDy(Dy, ValCix)
Dim DicId  As Dictionary:  Set DicId = DiKqNum(Col)
Dim DicCnt As Dictionary: Set DicCnt = DiKqCnt(Col)
Dim NCol%:                      NCol = NColzDy(Dy)
Dim ODy():                       ODy = Dy
Dim Dr, R&: For Each Dr In Itr(Dy)
:              ReDim Preserve Dr(NCol + 1) ' Extend 2 elements
    Dim K: K = Dr(ValCix)
    Dr(NCol) = DicId(K)
Dr(NCol + 1) = DicCnt(K)
      ODy(R) = Dr                          ' <== Put to ODy
           R = R + 1
Next
AddColzValIdqCnt_ToDy = ODy
End Function


