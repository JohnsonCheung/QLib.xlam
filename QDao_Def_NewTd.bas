Attribute VB_Name = "QDao_Def_NewTd"
Option Compare Text
Option Explicit
Private Const Asm$ = "QDao"
Private Const CMod$ = "MDao_Def_Td_New."

Private Sub AddPk(A As Dao.TableDef)
'Any Pk Fields in A.Fields?, if no exit sub
Dim F As Dao.Field2, IdFldNm$, J%
IdFldNm = A.Name & "Id"
If IsFdId(A.Fields(0), A.Name) Then
    A.Indexes.Append PkizT(A.Name)
    Exit Sub
End If
For J = 2 To A.Fields.Count
    If A.Fields(J).Name = IdFldNm Then Thw CSub, "The Table Id fields must be the fst fld", "I-th", J
Next
End Sub

Private Sub AddSk(A As Dao.TableDef, Skff$)
Dim SkFny$(): SkFny = TermAy(Skff): If Si(SkFny) = 0 Then Exit Sub
A.Indexes.Append NewSkIdx(A, SkFny)
End Sub

Private Function CvIdxFds(A) As Dao.IndexFields
Set CvIdxFds = A
End Function

Private Function IsFdId(A As Dao.Field2, T) As Boolean
If A.Name <> T & "Id" Then Exit Function
If A.Attributes <> Dao.FieldAttributeEnum.dbAutoIncrField Then Exit Function
If A.Type <> dbLong Then Exit Function
IsFdId = True
End Function

Function NewSkIdx(T As Dao.TableDef, SkFny$()) As Dao.Index
Const CSub$ = CMod & "NewSkIdx"
Dim O As New Dao.Index
O.Name = "SecondaryKey"
O.Unique = True
If Not HasEleAy(FnyzTd(T), SkFny) Then
    Thw CSub, "Given Td does not contain all given-SkFny", "Missing-SkFny Td-Name Td-Fny Given-SkFny", T.Name & "Id", MinusAy(SkFny, FnyzTd(T)), T.Name, FnyzTd(T), SkFny
End If
Dim IdxFds As Dao.IndexFields, I
Set IdxFds = CvIdxFds(O.Fields)
For Each I In SkFny
    IdxFds.Append Fd(CStr(I))
Next
Set NewSkIdx = O
End Function

Private Function PkizT(T) As Dao.Index
Dim O As New Dao.Index
O.Name = "PrimaryKey"
O.Primary = True
CvIdxFds(O.Fields).Append FdzId(T & "Id")
Set PkizT = O
End Function

Function TdzNm(T) As Dao.TableDef
Set TdzNm = New TableDef
TdzNm.Name = T
End Function
