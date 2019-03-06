Attribute VB_Name = "MDta_ExpLines"
Option Explicit

Function DrExpLinesCol(Dr, LinesColIx%) As Variant()
Dim A$()
    A = SplitCrLf(CStr(Dr(LinesColIx)))
Dim O()
    Dim IDr
        IDr = Dr
    Dim I
    For Each I In A
        IDr(LinesColIx) = I
        Push O, IDr
    Next
DrExpLinesCol = O
End Function

Function DrsExpLinesCol(A As Drs, LinesColNm$) As Drs
Dim Dry(): Dry = A.Dry
If Sz(Dry) = 0 Then
    Set DrsExpLinesCol = Drs(A.Fny, Dry)
    Exit Function
End If
Dim Ix%
    Ix = IxzAy(A.Fny, LinesColNm)
Dim O()
    Dim Dr
    For Each Dr In Dry
        PushAy Dry, DrExpLinesCol(Dr, Ix)
    Next
Set DrsExpLinesCol = Drs(A.Fny, O)
End Function
