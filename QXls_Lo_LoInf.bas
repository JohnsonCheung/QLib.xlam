Attribute VB_Name = "QXls_Lo_LoInf"
Private Const Asm$ = "Q"
Private Const CMod$ = "MLo."
Function NRowzLo&(A As ListObject)
NRowzLo = A.DataBodyRange.Rows.Count
End Function

Function LoNm$(T)
LoNm = "T_" & RmvFstNonLetter(T)
End Function

Function CvLo(A) As ListObject
Set CvLo = A
End Function

Function LoAllCol(A As ListObject) As Range
Set LoAllCol = RgzLoCC(A, 1, LoNCol(A))
End Function

Function LoAllEntCol(A As ListObject) As Range
Set LoAllEntCol = LoAllCol(A).EntireColumn
End Function

Function RgzLc(A As ListObject, C, Optional InclTot As Boolean, Optional InclHdr As Boolean) As Range
Dim R As Range
Set R = A.ListColumns(C).DataBodyRange
If Not InclTot And Not InclHdr Then
    Set RgzLc = R
    Exit Function
End If
If InclTot Then Set RgzLc = RgzMoreBelow(R, 1)
If InclHdr Then Set RgzLc = RgzMoreTop(R, 1)
End Function


Function DrszLo(A As ListObject) As Drs
DrszLo = Drs(FnyzLo(A), DryLo(A))
End Function
Function DryLo(A As ListObject) As Variant()
DryLo = DryzSq(SqzLo(A))
End Function
Function DryRgColAy(Rg As Range, ColIxy) As Variant()
DryRgColAy = DryzSqCol(SqzRg(Rg), ColIxy)
End Function
Function DryRgzLoCC(Lo As ListObject, CC) As Variant() _
' Return as many column as columns in [CC] from Lo
DryRgzLoCC = DryRgColAy(Lo.DataBodyRange, Ixy(FnyzLo(Lo), CC))
End Function

Function DtaAdrzLo$(A As ListObject)
DtaAdrzLo = WsRgAdr(A.DataBodyRange)
End Function

Function EntColzLo(A As ListObject, C) As Range
Set EntColzLo = RgzLc(A, C).EntireColumn
End Function

Function RgzLoCC(A As ListObject, C1, C2, Optional InclTot As Boolean, Optional InclHdr As Boolean) As Range
Dim R1&, R2&, mC1%, mC2%
R1 = R1Lo(A, InclHdr)
R2 = R2Lo(A, InclTot)
mC1 = WsCnozLc(A, C1)
mC2 = WsCnozLc(A, C2)
Set RgzLoCC = WsRCRC(WszLo(A), R1, mC1, R2, mC2)
End Function

Function LozWsDta(A As Worksheet, Optional LoNm$) As ListObject
Set LozWsDta = CrtLozRg(RgzWs(A), LoNm)
End Function

Function FbtStrLo$(A As ListObject)
FbtStrLo = FbtStrQt(A.QueryTable)
End Function

Function FnyzLo(A As ListObject) As String()
FnyzLo = Itn(A.ListColumns)
End Function

Function HasLoC(Lo As ListObject, ColNm$) As Boolean
HasLoC = HasItn(Lo.ListColumns, ColNm)
End Function

Function IsLozNoDta(A As ListObject) As Boolean
IsLozNoDta = IsNothing(A.DataBodyRange)
End Function

Function HdrCellzLo(A As ListObject, FldNm) As Range
Dim Rg As Range: Set Rg = A.ListColumns(FldNm).Range
Set HdrCellzLo = RgRC(Rg, 1, 1)
End Function

Function LoNCol%(A As ListObject)
LoNCol = A.ListColumns.Count
End Function
Function LozWs(A As Worksheet, LoNm$) As ListObject 'Return LoOpt
Set LozWs = FstItmzNm(A.ListObjects, LoNm)
End Function

Function FstLo(A As Worksheet) As ListObject 'Return LoOpt
Set FstLo = FstItm(A.ListObjects)
End Function




