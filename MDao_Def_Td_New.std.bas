Attribute VB_Name = "MDao_Def_Td_New"
Option Explicit
Const CMod$ = "MDao_Td_New."
Function TdShtTySemiFldSsl(T, ShtTySemiFldSsl$) As DAO.TableDef

End Function
Private Function CvIdxFds(A) As DAO.IndexFields
Set CvIdxFds = A
End Function

Private Function FdIsId(A As DAO.Field2, T) As Boolean
If A.Name <> T & "Id" Then Exit Function
If A.Attributes <> DAO.FieldAttributeEnum.dbAutoIncrField Then Exit Function
If A.Type <> dbLong Then Exit Function
FdIsId = True
End Function

Function NewTdSTR(TdStr) As DAO.TableDef
'Set NewTdSTR = NewTdSTR_EF(TdStr, EmpEF)
End Function

Private Function NewIdxSK(T As DAO.TableDef, Sk$()) As DAO.Index
Const CSub$ = CMod & "NewIdxSK"
Dim O As New DAO.Index
'O.Name = C_SkNm
O.Unique = True
'If Not HasEleAy(TdFny(T), Sk) Then
    Thw CSub, "Given Td does not contain all given-Sk", "Missing-Sk Td-Name Td-Fny Given-Sk", T.Name & "Id", AyMinus(Sk, TdFny(T)), T.Name, TdFny(T), Sk
'End If
Dim I
For Each I In Sk
    CvIdxFds(O.Fields).Append NewFd(I)
Next
Set NewIdxSK = O
End Function

Function NewTd(T, FdAy() As DAO.Field, Optional SkFny0) As DAO.TableDef
Dim O As New DAO.TableDef, F
O.Name = T
For Each F In FdAy
    O.Fields.Append F
Next
TdAddIdxPK O ' add Pk
TdAddIdxSK O, SkFny0 ' add Sk
Set NewTd = O
End Function

Private Sub TdAddIdxPK(A As DAO.TableDef)
'Any Pk Fields in A.Fields?, if no exit sub
Dim F As DAO.Field2
For Each F In A.Fields
    If FdIsId(F, A.Name) Then
        A.Indexes.Append NewIdxPK(A.Name)
        Exit Sub
    End If
Next
End Sub

Function NewIdxPK(T) As DAO.Index
Dim O As New DAO.Index
O.Name = "PrimaryKey"
O.Primary = True
CvIdxFds(O.Fields).Append FdzId(T & "Id")
Set NewIdxPK = O
End Function

Private Sub TdAddIdxSK(A As DAO.TableDef, SkFny0)
Dim Sk$(): Sk = CvNy(SkFny0): If Sz(Sk) = 0 Then Exit Sub
A.Indexes.Append NewIdxSK(A, Sk)
End Sub

