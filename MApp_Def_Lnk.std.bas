Attribute VB_Name = "MApp_Def_Lnk"
Option Explicit
Const SchmLines$ = _
           "Tbl     $Lnk    *Id | InpNm     | FilTy Ffn Bexpr" _
& vbCrLf & "Tbl     $LnkFld     | InpNm Fld | ExtNm DaoMulTyStr" _
& vbCrLf & "Tbl     $LnkFilTy   | FilTy     | FilTyDes" _
& vbCrLf & "Fld*   *Id InpNm    | ExtNm DaoMulTyStr" _
& vbCrLf & "Fld    $LnkFld    *Id InpNm | ExtNm DaoMulTyStr" _
& vbCrLf & "TblVal $LnkFilTy 1 [aaaa]"

Sub EdtTblLnk()
With Access.Application
    .Visible = True
    .DoCmd.OpenTable "$Lnk"
End With
End Sub

Sub EnsLnkDef()
Static A$
If CDb.Name <> A Then
    'SchmEns Schm
End If
End Sub

Private Sub Z_EnsLnkDef()
EnsLnkDef
End Sub

Private Sub Z()
Z_EnsLnkDef
MApp_Def_Lnk:
End Sub
