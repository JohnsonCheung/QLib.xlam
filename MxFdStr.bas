Attribute VB_Name = "MxFdStr"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxFdStr."
Public Const DaoTynn$ = "Boolean Byte Integer Int Long Single Double Char Text Memo Attachment" ' used in TzPFld
Sub CrtTblzPFld(D As Database, T, FdCsvLin$)
'Fm PrimNN: #Sql-FldLis-Phrase.  !The fld spec of create table sql inside the bkt.  Each fld sep by comma.  The spec allows:
'                                 !Boolean Byte Integer Int Long Single Double Char Text Memo Attachment
'Ret : create the @T in @D by DAO @@
Dim Td As DAO.TableDef: Set Td = NewTd(T)
AddFdy Td, FdAy(FdCsvLin)
D.TableDefs.Append Td
End Sub
Function FdStr$(F As DAO.Field2)
FdStr = F.Name & " " & DaoTyStr(F.Type) & IIf(F.Type = DAO.DataTypeEnum.dbText, " " & F.Size, "")
End Function

Function FfdStr$(A As DAO.Field2)
Dim D$, R$, Z$, VTxt$, VRul, E$, S$
If A.Type = DAO.DataTypeEnum.dbText Then S = " TxtSz=" & A.Size
If A.DefaultValue <> "" Then D = "Dft=" & A.DefaultValue
If A.Required Then R = "Req"
If A.AllowZeroLength Then Z = "AlwZLen"
If A.Expression <> "" Then E = "Expr=" & A.Expression
If A.ValidationRule <> "" Then VRul = "VRul=" & A.ValidationRule
If A.ValidationText <> "" Then VTxt = "VTxt=" & A.ValidationText
FfdStr = TLinzAp(A.Name, ShtDaoTy(A.Type), R, Z, VTxt, VRul, D, E, IIf((A.Attributes And DAO.FieldAttributeEnum.dbAutoIncrField) <> 0, "Auto", ""))
End Function

Function FdzS(FdStr) As DAO.Field2
Dim N$, S$ ' #Fldn and #Spec
Dim O As DAO.Field2
AsgBrkSpc FdStr, N, S
Select Case True
Case S = "Boolean":  Set O = FdzBool(N)
Case S = "Byte":     Set O = FdzByt(N)
Case S = "Integer", S = "Int": Set O = FdzInt(N)
Case S = "Long":     Set O = FdzLng(N)
Case S = "Single":   Set O = FdzSng(N)
Case S = "Double":   Set O = FdzDbl(N)
Case S = "Currency": Set O = FdzCur(N)
Case S = "Char":     Set O = FdzChr(N)
Case HasPfx(S, "Text"): Set O = FdzTxt(N, BetBkt(S))
Case S = "Memo":     Set O = FdzMem(N)
Case S = "Attachment": Set O = FdzAtt(N)
Case S = "Time":     Set O = FdzTim(N)
Case S = "Date":     Set O = FdzDte(N)
Case Else: Thw CSub, "Invalid FdStr", "Nm Spec vdt-DaoTynn, N, S, DaoTynn"
End Select
Set FdzS = O
End Function

Function FdAy(FdCsvLin$) As DAO.Field2()
':PFldCsv: :CsvLin #Sql-Fld-Phrase-Csv#  ! Each Itm is :PFld. It uses DAO to create
Dim FdStr: For Each FdStr In Itr(AmTrim(SplitComma(FdCsvLin)))
    PushObj FdAy, FdzS(FdStr)
Next
End Function


