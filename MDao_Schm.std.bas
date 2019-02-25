Attribute VB_Name = "MDao_Schm"
Option Explicit
Const CMod$ = "MDao_Schm."
Const C_Tbl$ = "Tbl"
Const C_Fld$ = "Fld"
Const C_Ele$ = "Ele"
Const C_Des$ = "Des"
Type EFDic
    EleToEdStrDic As Dictionary
    FDicSmf As Dictionary
End Type

Function FnySchm(Schm$()) As String()

End Function
Sub CrtSchmz(A As Database, Schm$())
Const CSub$ = CMod & "CrtSchmz"
ThwErMsg ErSchm(Schm), CSub, "there is error in the Schm", "Schm Db", AyAddIxPfx(Schm, 1), DbNm(A)
Dim Smt$(): Smt = AywRmvT1(Schm, C_Tbl)
Dim EDic As Dictionary
Dim FDic As Dictionary
    Set FDic = KeyToLikAyDic_T1LikssLy(AywRmvT1(Schm, C_Fld))
    Set EDic = Dic(AywRmvT1(Schm, C_Ele))
AppTdAyz A, TdAy(Smt, FDic, EDic)
RunSqyz A, SqyCrtPk_Tny(PkTnySmt(Smt))
RunSqyz A, SqyCrtSkTdStrAy(Smt)
Set TblDesDicz(A) = Dic(AywRmvTT(Schm, C_Des, C_Tbl))
Set FldDesDicz(A) = Dic(AywRmvTT(Schm, C_Des, C_Fld))
End Sub

Function HasDbtf(A As Database, T, F) As Boolean
If Not HasTblz(A, T) Then Exit Function
HasDbtf = HasItn(A.TableDefs(T).Fields, F)
End Function

Private Function PkTnySmt(Smt$()) As String()
Dim L
For Each L In Smt
    If HasSubStr(L, " *Id ") Then
        PushI PkTnySmt, T1(L)
    End If
Next
End Function
Private Function SqyCrtSkTdStrAy(A$()) As String()
Dim TdStr, SkFny$()
For Each TdStr In Itr(A)
    SkFny = SkFnyTdStr(TdStr)
    If Sz(SkFny) > 0 Then
'        PushI SqyCrtSkTdStrAy, SqlCrtSk(T1(TdStr), SkFny)
    End If
Next
End Function

Private Function SkFnyTdStr(TdStr) As String()
Dim P%, T$, Rst$
P = InStr(TdStr, "|")
If P = 0 Then Exit Function
'AsgTRst TakBef(Smt, "|"), T, Rst
Rst = Replace(Rst, T, "*")
SkFnyTdStr = SySsl(Rst)
End Function

Private Function TdAy(TdStrAy$(), FDic As Dictionary, EDic As Dictionary) As DAO.TableDef()
Dim T
For Each T In T1Ay(TdStrAy)
    PushObj TdAy, Td(T, FDic, EDic)
Next
End Function
Private Function HasPkTdStr(TdStr) As Boolean
HasPkTdStr = HasSubStr(TdStr, " *Id ")
End Function

Private Function Td(TdStr, FDic As Dictionary, EDic As Dictionary) As DAO.TableDef
'FDic is FDicSmf
Dim O As DAO.TableDef
Dim IsPk As Boolean
Dim F
For Each F In FnySmt(TdStr)
    If HasPkTdStr(TdStr) Then
        O.Fields.Append FdzPk(F)
    Else
        O.Fields.Append Fd(F, FDic, EDic)
    End If
Next
O.Name = T1(TdStr)
End Function
Private Function FdStrFldFDicEDic$(Fld, FDic As Dictionary, EDic As Dictionary)
Dim Ele$
Ele = EleFldFDic(Fld, FDic)
FdStrFldFDicEDic = FdStrFldEleEDic(Fld, Ele, EDic)
End Function
Private Function Fd(Fld, FDic As Dictionary, EDic As Dictionary) As DAO.Field
Set Fd = FdzFdStr(FdStrFldFDicEDic(Fld, FDic, EDic))
End Function

Private Function EleStrEleEDic$(Ele, EDic As Dictionary)

End Function
Private Function FdStrFldEleEDic$(Fld, Ele$, EDic As Dictionary)
FdStrFldEleEDic = Fld & " " & EleStrEleEDic(Ele, EDic)
End Function
Private Function EleFldFDic$(Fld, FDic As Dictionary)
EleFldFDic = Keyz_LikAyDic_Itm(FDic, Fld)
End Function
Private Function FnySmt(A) As String()

End Function

Private Sub Z_CrtSchmz()
Dim Schm$(), Db As Database
Set Db = TmpDb
Schm = SampSchm
GoSub Tst
Exit Sub
Tst:
    CrtSchmz Db, Schm
'    Debug.Assert HasDbt(Db, "A")
'    Debug.Assert HasDbt(Db, "B")
'    Debug.Assert IsEqAy(FnyDbt(Db, "A"), SySsl("AId ANm ADte AATy Loc Expr Rmk"))
'    Debug.Assert IsEqAy(FnyDbt(Db, "B"), SySsl("BId ANm BNm BDte"))
    Brw Db
    Stop
    Return
End Sub

Private Property Get Z_CrtSchm1$()
Const A_1$ = "Tbl A *Id | *Nm     | *Dte AATy Loc Expr Rmk" & _
vbCrLf & "Fld B *Id | AId *Nm | *Dte" & _
vbCrLf & "Fld Txt AATy" & _
vbCrLf & "Fld Mem Rmk" & _
vbCrLf & "Ele Loc Txt Rq Dft=ABC [VTxt=Loc must cannot be blank] [VRul=IsNull([Loc]) or Trim(Loc)='']" & _
vbCrLf & "Ele Expr Txt [Expr=Loc & 'abc']" & _
vbCrLf & "Des Tbl  A     AA BB " & _
vbCrLf & "Des Tbl  A     CC DD " & _
vbCrLf & "Des Fld  ANm   AA BB " & _
vbCrLf & "Des Fld  A.ANm TF_Des-AA-BB"

Z_CrtSchm1 = A_1
End Property

Private Sub Z_CrtSchm2()
Dim Td As DAO.TableDef
Dim B As DAO.Field2
GoSub X_Td
GoSub Tst
Exit Sub
Tst:
    Debug.Print ObjPtr(Td)
    CDb.TableDefs.Append Td
    Debug.Print ObjPtr(Td)
    Set B = CDb.TableDefs("#Tmp").Fields("B")
    CDb.TableDefs("#Tmp").Fields("B").Properties.Append CDb.CreateProperty(C_Des, dbText, "ABC")
    Return
X_Td:
    Dim FdAy() As DAO.Field
    Set B = NewFd("B", dbInteger)
    PushObj FdAy, FdzId("#Tmp")
    PushObj FdAy, NewFd("A", dbInteger)
    PushObj FdAy, B
    Set Td = NewTd("#Tmp", FdAy, "A B")
    Return
End Sub

Private Sub ZZ()
End Sub
Sub AA()
Z
End Sub
Private Sub Z()
Z_CrtSchmz
End Sub
