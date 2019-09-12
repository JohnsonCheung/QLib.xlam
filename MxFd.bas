Attribute VB_Name = "MxFd"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxFd."

Function CvFd(A) As dao.Field
Set CvFd = A
End Function

Function CvFd2(A As dao.Field) As dao.Field2
Set CvFd2 = A
End Function

Function FdClone(A As dao.Field2, FldNm) As dao.Field2
Set FdClone = New dao.Field
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

Function FdStr$(A As dao.Field2)
Dim D$, R$, Z$, VTxt$, VRul, E$, S$
If A.Type = dao.DataTypeEnum.dbText Then S = " TxtSz=" & A.Size
If A.DefaultValue <> "" Then D = "Dft=" & A.DefaultValue
If A.Required Then R = "Req"
If A.AllowZeroLength Then Z = "AlwZLen"
If A.Expression <> "" Then E = "Expr=" & A.Expression
If A.ValidationRule <> "" Then VRul = "VRul=" & A.ValidationRule
If A.ValidationText <> "" Then VTxt = "VTxt=" & A.ValidationText
FdStr = TLinzAp(A.Name, ShtTyzDao(A.Type), R, Z, VTxt, VRul, D, E, IIf((A.Attributes And dao.FieldAttributeEnum.dbAutoIncrField) <> 0, "Auto", ""))
End Function

Function FdStrAyFds(A As dao.Fields) As String()
Dim F As dao.Field
For Each F In A
    PushI FdStrAyFds, FdStr(F)
Next
End Function

Function IsEqFd(A As dao.Field2, B As dao.Field2) As Boolean
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

Function Fv(A As dao.Field)
On Error Resume Next
Fv = A.Value
End Function
