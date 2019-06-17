Attribute VB_Name = "QIde_B_LinInf"
Option Explicit
Option Compare Text
Function IsRmkzLin(L) As Boolean
IsRmkzLin = FstChr(LTrim(L)) = "'"
End Function
