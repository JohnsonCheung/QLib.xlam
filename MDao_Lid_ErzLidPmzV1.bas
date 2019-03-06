Attribute VB_Name = "MDao_Lid_ErzLidPmzV1"
Dim A As LidPm
Private Type TycOpt
    Tyc As LidMisTyc
    Som As Boolean
End Type
Private Type Tbl
    '1
    Ffn As String
    FilNm As String
    TblNm As String
    Wsn As String
    EptFset As Aset
    FldNmToEptShtTyLisDic As Dictionary
    '2
    IsFilExist As Boolean
    '3
    IsTblExist As Boolean
    '4
    ActFset As Aset
    '5
    Tyc() As LidMisTyc
End Type

Function ErzLidPmzV1(LidPm As LidPm) As String()
Set A = LidPm
Dim FfnAy$():  FfnAy = FfnAyzLidFil(A.Fil)
Dim Exist$():  Exist = ExistFfnAy(FfnAy)
Dim F As Aset: Set F = MisFfnAset(FfnAy)
Dim T1() As Tbl:  T1 = T1Ay(A)
Dim T2() As Tbl:  T2 = T2Ay(T1, Exist)
Dim T3() As Tbl:  T3 = T3Ay(T2)
Dim T4() As Tbl:  T4 = T4Ay(T3)
Dim T5() As Tbl:  T5 = T5Ay(T4)
Dim OFfn$(): OFfn = MsgzMisFfn(F)
Dim OTbl$(): OTbl = MsgzMisTbl(T5)
Dim OCol$(): OCol = MsgzMisCol(T5)
Dim OTy$():   OTy = MsgzMisTy(T5)
ErzLidPmzV1 = SyAddAp(OFfn, OTbl, OCol, OTy)
End Function

Private Function MsgzMisFfn(MisFfn As Aset, Optional FilKind$ = "file") As String()
If MisFfn.IsEmp Then Exit Function
Dim L1$
    Dim S$
    If MisFfn.Cnt = 1 Then
        S = FmtQQ("one ?", FilKind)
    Else
        S = FmtQQ("? ?s", MisFfn.Cnt, FilKind)
    End If
    L1 = FmtQQ("Following ? not found", S)
PushI MsgzMisFfnToFilKdDic, L1
PushIAy MsgzMisFfnToFilKdDic, AyTab(LyzGpPth(MisFfn))
End Function

Private Function MsgzMisFfnAy(FfnAy$(), Optional FilKind$ = "file") As String()
MsgzMisFfnAy = MsgzMisFfn(AsetzAy(FfnAy), FilKind)
End Function

Private Function MsgzMisTbl(T() As Tbl) As String()
Dim I%, M As LidMisTbl
For I = 0 To UBound(T)
    With T(I)
        Set M = New LidMisTbl
        PushIAy MsgzMisTbl, MsgzMisTbl1(T(I))
    End With
Next
End Function
Private Function MsgzMisTbl1(A As Tbl) As String()

End Function
'======================================================================
'======================================================================
'======================================================================
Private Function MsgzMisCol(T() As Tbl) As String()
Dim I%
For I = 0 To UBound(T)
    With T(I)
        If Not .IsTblExist Then GoTo Nxt
        If .EptFset.Minus(.ActFset).Cnt > 0 Then
            Set M = New LidMisCol
            PushIAy MsgzMisCol, MsgzMisCol1(T(I))
        End If
    End With
Nxt:
Next
End Function
Private Function MsgzMisCol1(T As Tbl) As String()
With T
    If Not .IsTblExist Then Exit Function
    If .EptFset.Minus(.ActFset).Cnt > 0 Then Exit Function
    PushIAy MsgzMisCol1, MsgzMisCol2(T)
End With
End Function
Private Function MsgzMisCol2(T As Tbl) As String()

End Function
'======================================================================
'======================================================================
'======================================================================

Private Function MsgzMisTy(T() As Tbl) As String()
Dim I%, M As LidMisTy
For I = 0 To UBound(T)
    With T(I)
        If Sz(.Tyc) > 0 Then Exit Function
        PushIAy MsgzMisTy, MsgzMisTy1(T(I))
    End With
Next
End Function
Private Function MsgzMisTy1(T As Tbl) As String()

End Function
'=============================================================
'=============================================================
'=============================================================
Private Function T1Ay(A As LidPm) As Tbl()
Dim U%, UFx%, UFb%
    UFx = UB(A.Fx)
    UFb = UB(A.Fb)
    U = UFx + UFb + 1
   
Dim O() As Tbl
    ReDim O(U)
Dim J%, I, MFx As LidFx, MFb As LidFb
Dim D As Dictionary: Set D = FfnDic(A.Fil)
J = 0
For Each I In Itr(A.Fx)
    Set MFx = I
    O(J) = T1Fx(MFx, A.Fil, D)
    J = J + 1
Next
For Each I In Itr(A.Fb)
    Set MFb = I
    O(J) = T1Fb(MFb, A.Fil)
    J = J + 1
Next
T1Ay = O
End Function
Private Function FfnDic(A() As LidFil) As Dictionary
Set FfnDic = New Dictionary
Dim I, M As LidFil
For Each I In A
    Set M = I
    FfnDic.Add M.FilNm, M.Ffn
Next
End Function
Private Function T1Fx(A As LidFx, B() As LidFil, FfnDic As Dictionary) As Tbl
With T1Fx
    .Ffn = FfnDic(A.Fxn)
    .FilNm = A.Fxn
    Set .EptFset = FsetzFxc(A.Fxc)
    .Wsn = A.Wsn
    Set .FldNmToEptShtTyLisDic = FldNmToEptShtTyLisDiczFxc(A.Fxc)
End With
End Function
Private Function FsetzFxc(A() As LidFxc) As Aset
Set FsetzFxc = New Aset
Dim I, M As LidFxc
For Each I In A
    Set M = I
    FsetzFxc.PushItm M.ExtNm
Next
End Function
Private Function FldNmToEptShtTyLisDiczFxc(A() As LidFxc) As Dictionary
Dim I, M As LidFxc
Set FldNmToEptShtTyLisDiczFxc = New Dictionary
For Each I In A
    Set M = I
    FldNmToEptShtTyLisDiczFxc.Add M.ExtNm, M.ShtTyLis
Next
End Function

Private Function T1Fb(A As LidFb, B() As LidFil) As Tbl
With T1Fb
    .Ffn = A.Fb
    .FilNm = A.Fbn
    .EptFset = A.Fset
End With
End Function

'-----------------------------------------------------------------------------------
Private Function T2Ay(T() As Tbl, ExistFfnAy$()) As Tbl()
Dim J%
For J = 0 To UBound(T)
    T(J).IsFilExist = HasEle(ExistFfnAy, T(J).Ffn)
Next
T2Ay = T
End Function
Private Function T3Ay(T() As Tbl) As Tbl()
Dim J%
For J = 0 To UBound(T)
    With T(J)
        If .IsFilExist Then
            .IsTblExist = HasTblzFfnTblNm(.Ffn, .TblNm)
        End If
    End With
Next
T3Ay = T
End Function
Private Function T4Ay(T() As Tbl) As Tbl()
Dim J%
For J = 0 To UBound(T)
    With T(J)
        If .IsTblExist Then
            Set .ActFset = AsetzAy(FnyzFfnTblNm(.Ffn, .TblNm))
        End If
    End With
Next
T4Ay = T
End Function

Private Function T5Ay(T() As Tbl) As Tbl()
Dim J%, ExistFset As Aset
For J = 0 To UBound(T)
    With T(J)
        If .IsTblExist Then
            Set ExistFset = .EptFset.Minus(.ActFset)
            If ExistFset.Cnt > 0 Then
                .Tyc = Tyc(ExistFset, .FldNmToEptShtTyLisDic, .Ffn, .TblNm)
            End If
            Set .ActFset = AsetzAy(FnyzFfnTblNm(.Ffn, .TblNm))
        End If
    End With
Next
T5Ay = T
End Function

Private Function Tyc(ExistFset As Aset, FldNmToEptShtTyLisDic As Dictionary, Ffn$, TblNm$) As LidMisTyc()
Dim F
Dim ActShtTy$, EptShtTyLis$
Dim ActDic As Dictionary: Set ActDic = ShtTyDic(Ffn, TblNm)

For Each F In ExistFset.Itms
    ActShtTy = ActDic(F)
    EptShtTyLis = FldNmToEptShtTyLisDic(F)
    With Tyci(ActShtTy, EptShtTyLis, F)
        If .Som Then
            PushObj Tyc, .Tyc
        End If
    End With
Next
End Function

Private Function Tyci(ActShtTy$, EptShtTyLis$, ExtNm) As TycOpt
If HasEle(CmlAy(EptShtTyLis), ActShtTy) Then Exit Function
Set Tyci.Tyc = New LidMisTyc
Tyci.Tyc.Init ExtNm, ActShtTy, EptShtTyLis
Tyci.Som = True
End Function

'===================================================================================
Private Function FfnzFilNm$(FilNm$, A() As LidFil)
Dim J%
For J = 0 To UB(A)
    If A(J).FilNm = FilNm Then
        FfnzFilNm = A(J).Ffn
    End If
Next
End Function

Private Function FfnAyzLidFil(A() As LidFil) As String()
Dim J%
For J = 0 To UB(A)
    PushI FfnAyzLidFil, A(J).Ffn
Next
End Function


