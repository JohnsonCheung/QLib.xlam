Attribute VB_Name = "QDao_Def_Idx"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Idx."
Private Const Asm$ = "QDao"

Function CvIdx(A) As Dao.Index
Set CvIdx = A
End Function

Function FnyzIdx(A As Dao.Index) As String()
If IsNothing(A) Then Exit Function
FnyzIdx = Itn(A.Fields)
End Function

Function IdxStr$(A As Dao.Index)
Dim X$, F$
With A
IdxStr = FmtQQ("Idx;?;?;?", .Name, X, F)
End With
End Function

Function IdxStrAyIdxs(A As Dao.Indexes) As String()
Dim I As Dao.Index
For Each I In A
    PushI IdxStrAyIdxs, IdxStr(I)
Next
End Function

Function IsEqIdx(A As Dao.Index, B As Dao.Index) As Boolean
With A
Select Case True
Case .Name <> B.Name
Case .Primary <> B.Primary
Case .Unique <> B.Unique
Case Not IsEqAy(Itn(.Fields), Itn(B.Fields))
Case Else: IsEqIdx = True
End Select
End With
End Function

Function IsEqIdxs(A As Dao.Indexes, B As Dao.Indexes) As Boolean
If A.Count <> B.Count Then Exit Function
If Not IsEqNmItr(A, B) Then Exit Function
Dim I
For Each I In A
    If Not IsEqIdx(CvIdx(I), B(CvIdx(I).Name)) Then Exit Function
Next
End Function

Function IsIdxSk(A As Dao.Index, T) As Boolean
If A.Name <> T Then Exit Function
IsIdxSk = A.Unique
End Function

Function IsIdxUniq(A As Dao.Index) As Boolean
If IsNothing(A) Then Exit Function
IsIdxUniq = A.Unique
End Function
