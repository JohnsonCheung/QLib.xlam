Attribute VB_Name = "MxMdSts"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxMdSts."
Type MdSts
    NLin As Long
    NPub As Integer
    NPrv As Integer
    NFrd As Integer
End Type

Function MdSts(M As CodeModule) As MdSts
Dim S$(): S = Src(M)
Dim Mth$(): Mth = MthLinAy(S)
With MdSts
    .NLin = Si(S)
    Dim L: For Each L In Itr(Mth)
        Select Case MthMdy(L)
        Case "", "Public": .NPub = .NPub + 1
        Case "Private":    .NPrv = .NPrv + 1
        Case "Friend":     .NFrd = .NFrd + 1
        Case Else: Stop
        End Select
    Next
End With
End Function

