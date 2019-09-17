Attribute VB_Name = "MxCrtLo"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxCrtLo."

Function CrtLoAtzSq(Sq(), At As Range, Optional Lon$) As ListObject
Set CrtLoAtzSq = CrtLo(RgzSq(Sq(), At), Lon)
End Function

Function CrtLo(Rg As Range, Optional Lon$) As ListObject
Dim S As Worksheet: Set S = WszRg(Rg)
Dim O As ListObject: Set O = S.ListObjects.Add(xlSrcRange, Rg, , xlYes)
BdrAround Rg
Rg.EntireColumn.AutoFit
SetLon O, Lon
Set CrtLo = O
End Function

Function CrtLoAtzDrs(D As Drs, At As Range, Optional Lon$) As ListObject
Set CrtLoAtzDrs = CrtLo(RgzDrs(D, At), Lon)
End Function

Function CrtEmpLo(At As Range, FF$, Optional Lon$) As ListObject
Set CrtEmpLo = CrtLo(RgzAyH(SyzSS(FF), At), Lon)
End Function


