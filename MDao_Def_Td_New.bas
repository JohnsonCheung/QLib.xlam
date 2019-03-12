Attribute VB_Name = "MDao_Def_Td_New"
Option Explicit
Const CMod$ = "MDao_Td_New."

Private Function CvIdxFds(A) As Dao.IndexFields
Set CvIdxFds = A
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
    CvIdxFds(O.Fields).Append Fd(I)
Next
Set SkIdx = O
End Function

Function TdzFdAy(T, FdAy() As Field2, Optional SkFF) As Dao.TableDef
Dim O As New Dao.TableDef, F
O.Name = T
AppFdAy O, FdAy
AddPk O
AddSk O, SkFF
Set TdzFdAy = O
End Function

Private Sub AddPk(A As Dao.TableDef)
'Any Pk Fields in A.Fields?, if no exit sub
Dim F As Dao.Field2
For Each F In A.Fields
    If IsIdFd(F, A.Name) Then
        A.Indexes.Append PkIdxzT(A.Name)
        Exit Sub
    End If
Next
End Sub

Function PkIdxzT(T) As Dao.Index
Dim O As New Dao.Index
O.Name = "PrimaryKey"
O.Primary = True
CvIdxFds(O.Fields).Append FdzId(T & "Id")
Set PkIdxzT = O
End Function

Private Sub AddSk(A As Dao.TableDef, SkFF)
Dim Sk$(): Sk = NyzNN(SkFF): If Sz(Sk) = 0 Then Exit Sub
A.Indexes.Append SkIdx(A, Sk)
End Sub

