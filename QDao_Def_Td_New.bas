Attribute VB_Name = "QDao_Def_Td_New"
Option Compare Text
Option Explicit
Private Const Asm$ = "QDao"
Private Const CMod$ = "MDao_Def_Td_New."

Private Function CvIdxFds(A) As DAO.IndexFields
Set CvIdxFds = A
End Function

Private Function IsFdId(A As DAO.Field2, T) As Boolean
If A.Name <> T & "Id" Then Exit Function
If A.Attributes <> DAO.FieldAttributeEnum.dbAutoIncrField Then Exit Function
If A.Type <> dbLong Then Exit Function
IsFdId = True
End Function

Function NewSkIdx(T As DAO.TableDef, SkFny$()) As DAO.Index
Const CSub$ = CMod & "NewSkIdx"
Dim O As New DAO.Index
O.Name = "SecondaryKey"
O.Unique = True
If Not HasEleAy(FnyzTd(T), SkFny) Then
    Thw CSub, "Given Td does not contain all given-SkFny", "Missing-SkFny Td-Name Td-Fny Given-SkFny", T.Name & "Id", MinusAy(SkFny, FnyzTd(T)), T.Name, FnyzTd(T), SkFny
End If
Dim IdxFds As DAO.IndexFields, I
Set IdxFds = CvIdxFds(O.Fields)
For Each I In SkFny
    IdxFds.Append Fd(CStr(I))
Next
Set NewSkIdx = O
End Function

Function TdzTF(T, Fdy() As DAO.Field2, Optional Skff$) As DAO.TableDef
Dim O As New TableDef, F
O.Name = T
AddSk O, Skff
AddPk O
AddFdy O, Fdy
Set TdzTF = O
End Function

Private Sub AddPk(A As DAO.TableDef)
'Any Pk Fields in A.Fields?, if no exit sub
Dim F As DAO.Field2, IdFldNm$, J%
IdFldNm = A.Name & "Id"
If IsFdId(A.Fields(0), A.Name) Then
    A.Indexes.Append PkizT(A.Name)
    Exit Sub
End If
For J = 2 To A.Fields.Count
    If A.Fields(J).Name = IdFldNm Then Thw CSub, "The Table Id fields must be the fst fld", "I-th", J
Next
End Sub

Private Function PkizT(T) As DAO.Index
Dim O As New DAO.Index
O.Name = "PrimaryKey"
O.Primary = True
CvIdxFds(O.Fields).Append FdzId(T & "Id")
Set PkizT = O
End Function

Private Sub AddSk(A As DAO.TableDef, Skff$)
Dim SkFny$(): SkFny = TermAy(Skff): If Si(SkFny) = 0 Then Exit Sub
A.Indexes.Append NewSkIdx(A, SkFny)
End Sub

