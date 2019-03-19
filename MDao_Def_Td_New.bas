Attribute VB_Name = "MDao_Def_Td_New"
Option Explicit
Const CMod$ = "MDao_Td_New."

Private Function CvIdxfds(A) As Dao.IndexFields
Set CvIdxfds = A
End Function

Private Function IsIdFd(A As Dao.Field2, T) As Boolean
If A.Name <> T & "Id" Then Exit Function
If A.Attributes <> Dao.FieldAttributeEnum.dbAutoIncrField Then Exit Function
If A.Type <> dbLong Then Exit Function
IsIdFd = True
End Function

Private Function SkIdx(T As Dao.TableDef, Sk$()) As Dao.Index
Const CSub$ = CMod & "SkIdx"
Dim O As New Dao.Index
O.Name = "SecondaryKey"
O.Unique = True
If Not HasEleAy(FnyzTd(T), Sk) Then
    Thw CSub, "Given Td does not contain all given-Sk", "Missing-Sk Td-Name Td-Fny Given-Sk", T.Name & "Id", AyMinus(Sk, FnyzTd(T)), T.Name, FnyzTd(T), Sk
End If
Dim I
For Each I In Sk
    CvIdxfds(O.Fields).Append Fd(I)
Next
Set SkIdx = O
End Function

Function TdzFdy(T, Fdy() As Dao.Field2, Optional SkFF) As Dao.TableDef
Dim O As New Dao.TableDef, F
O.Name = T
Set TdzFdy = TdAddSk(TdAddPk(TdAppFdy(O, Fdy)), SkFF)
End Function

Private Function TdAddPk(A As Dao.TableDef) As Dao.TableDef
'Any Pk Fields in A.Fields?, if no exit sub
Dim F As Dao.Field2
For Each F In A.Fields
    If IsIdFd(F, A.Name) Then
        A.Indexes.Append PkIdxzT(A.Name)
        Exit Function
    End If
Next
Set TdAddPk = A
End Function

Function PkIdxzT(T) As Dao.Index
Dim O As New Dao.Index
O.Name = "PrimaryKey"
O.Primary = True
CvIdxfds(O.Fields).Append FdzId(T & "Id")
Set PkIdxzT = O
End Function

Private Function TdAddSk(A As Dao.TableDef, SkFF) As Dao.TableDef
Dim Sk$(): Sk = NyzNN(SkFF): If Si(Sk) = 0 Then Exit Function
A.Indexes.Append SkIdx(A, Sk)
Set TdAddSk = A
End Function

