Attribute VB_Name = "QDao_Att_Op_Imp"
Option Compare Text
Option Explicit
Private Const Asm$ = "QDao"
Private Const CMod$ = "MDao_Att_Op_Imp."

Private Sub ImpAttzAttd(A As Attd, Ffn$)
Const CSub$ = CMod & "ImpAttzAttd"
Dim F2 As Field2
Dim S&, T As Date
S = SizFfn(Ffn)
T = DtezFfn(Ffn)
'Msg CSub, "[Att] is going to import [Ffn] with [Si] and [Tim]", FdVal(A.TRs!AttNm), Ffn, S, T
With A
    .TRs.Edit
    With .Ars
        If HasReczFEv(A.Ars, "FileName", Fn(Ffn)) Then
            D "Ffn is found in Att and it is replaced"
            .Edit
        Else
            D "Ffn is not found in Att and it is imported"
            .AddNew
        End If
        Set F2 = !FileData
        F2.LoadFromFile Ffn
        .Update
    End With
    .TRs.Fields!FilTim = DtezFfn(Ffn)
    .TRs.Fields!FilSz = SizFfn(Ffn)
    .TRs.Update
End With
End Sub

Sub ImpAtt(A As Database, Att$, FmFfn$)
ImpAttzAttd Attd(A, Att), FmFfn
End Sub

Private Sub Z_ImpAtt()
Dim T$, D As Database
T = TmpFt
WrtStr "sdfdf", T
ImpAtt D, "AA", T
Kill T
'T = TmpFt
'ExpAttToFfn "AA", T
'BrwFt T
End Sub

Private Sub ZZ()
Z_ImpAtt
MDao_Att_Op_Imp:
End Sub
