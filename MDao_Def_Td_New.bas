Attribute VB_Name = "MDao_Def_Td_New"
Option Explicit
Const CMod$ = "MDao_Td_New."

Private Function CvIdxFds(A) As Dao.IndexFields
Set CvIdxFds = A
End Function

Private Function IsIdFd(A As Dao.Field2, T$) As Boolean
If A.Name <> T & "Id" Then Exit Function
If A.Attributes <> Dao.FieldAttributeEnum.dbAutoIncrField Then Exit Function
If A.Type <> dbLong Then Exit Function
IsIdFd = True
End Function

Function NewSkIdx(T As Dao.TableDef, SkFny$()) As Dao.Index
Const CSub$ = CMod & "NewSkIdx"
Dim O As New Dao.Index
O.Name = "SecondaryKey"
O.Unique = True
If Not HasEleAy(FnyzTd(T), SkFny) Then
    Thw CSub, "Given Td does not contain all given-SkFny", "Missing-SkFny Td-Name Td-Fny Given-SkFny", T.Name & "Id", AyMinus(SkFny, FnyzTd(T)), T.Name, FnyzTd(T), SkFny
End If
Dim IdxFds As Dao.IndexFields, I
Set IdxFds = CvIdxFds(O.Fields)
For Each I In SkFny
    IdxFds.Append Fd(CStr(I))
Next
Set NewSkIdx = O
End Function

Function TdzFdy(T$, Fdy() As Field2, Optional Skff$) As Dao.TableDef
Dim O As New Dao.TableDef, F
O.Name = T
AddSk O, Skff
AddPk O
AddFdy O, Fdy
Set TdzFdy = O
End Function

Private Sub AddPk(A As Dao.TableDef)
'Any Pk Fields in A.Fields?, if no exit sub
Dim F As Dao.Field2, IdFldNm$, J%
IdFldNm = A.Name & "Id"
If IsIdFd(A.Fields(0), A.Name) Then
    A.Indexes.Append NewPkIdx(A.Name)
    Exit Sub
End If
For J = 2 To A.Fields.Count
    If A.Fields(J).Name = IdFldNm Then Thw CSub, "The Table Id fields must be the fst fld", "I-th", J
Next
End Sub

Private Function NewPkIdx(T) As Dao.Index
Dim O As New Dao.Index
O.Name = "PrimaryKey"
O.Primary = True
CvIdxFds(O.Fields).Append FdzId(T & "Id")
Set NewPkIdx = O
End Function

Private Sub AddSk(A As Dao.TableDef, Skff$)
Dim SkFny$(): SkFny = TermAy(Skff): If Si(SkFny) = 0 Then Exit Sub
A.Indexes.Append NewSkIdx(A, SkFny)
End Sub

