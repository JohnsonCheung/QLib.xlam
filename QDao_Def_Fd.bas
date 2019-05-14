Attribute VB_Name = "QDao_Def_Fd"
Option Explicit
Private Const CMod$ = "MDao_Def_Fd."
Private Const Asm$ = "QDao"
Function FdClone(A As Dao.Field2, FldNm) As Dao.Field2
Set FdClone = New Dao.Field
With FdClone
    .Name = FldNm
    .Type = A.Type
    .AllowZeroLength = A.AllowZeroLength
    .Attributes = A.Attributes
    .DefaultValue = A.DefaultValue
    .Expression = A.Expression
    .Required = A.Required
    .ValidationRule = A.ValidationRule
    .ValidationText = A.ValidationText
End With
End Function

Function FdVal(A As Dao.Field)
FdVal = A.Value
End Function

Function IsEqFd(A As Dao.Field2, B As Dao.Field2) As Boolean
With A
    If .Name <> B.Name Then Exit Function
    If .Type <> B.Type Then Exit Function
    If .Required <> B.Required Then Exit Function
    If .AllowZeroLength <> B.AllowZeroLength Then Exit Function
    If .DefaultValue <> B.DefaultValue Then Exit Function
    If .ValidationRule <> B.ValidationRule Then Exit Function
    If .ValidationText <> B.ValidationText Then Exit Function
    If .Expression <> B.Expression Then Exit Function
    If .Attributes <> B.Attributes Then Exit Function
    If .Size <> B.Size Then Exit Function
End With
IsEqFd = True
End Function

Function CvFd(A) As Dao.Field
Set CvFd = A
End Function

Function CvFd2(A As Dao.Field) As Dao.Field2
Set CvFd2 = A
End Function
