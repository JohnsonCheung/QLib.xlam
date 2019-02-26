Attribute VB_Name = "MDao_Schm"
Option Explicit
Const CMod$ = "MDao_Schm."
Const C_Tbl$ = "Tbl"
Const C_Fld$ = "Fld"
Const C_Ele$ = "Ele"
Const C_Des$ = "Des"
Type EF
    EleLy() As String
    FldLy() As String
End Type

Sub CrtSchm(A As Database, Schm$())
Const CSub$ = CMod & "CrtSchm"
ThwErMsg ErzSchm(Schm), CSub, "there is error in the Schm", "Schm Db", AyAddIxPfx(Schm, 1), DbNm(A)
Dim TdLy$(): TdLy = AywRmvT1(Schm, C_Tbl)
Dim EF As EF
    EF = EFzSchm(Schm)
AppTdAyz A, TdAy(TdLy, EF)
RunSqy A, SqyCrtPkzTny(PkTny(TdLy))
RunSqy A, SqyCrtSk(TdLy)
Set TblDesDic(A) = Dic(AywRmvTT(Schm, C_Des, C_Tbl))
Set FldDesDic(A) = Dic(AywRmvTT(Schm, C_Des, C_Fld))
End Sub

Private Function EFzSchm(Schm$()) As EF
EFzSchm.EleLy = AywRmvT1(Schm, "Ele")
EFzSchm.FldLy = AywRmvT1(Schm, "Fld")
End Function

Private Function PkTny(TdLy$()) As String()
Dim L
For Each L In TdLy
    If HasSubStr(L, " *Id ") Then
        PushI PkTny, T1(L)
    End If
Next
End Function

Private Function SqyCrtSk(TdLy$()) As String()
Dim TdLin, Sk$()
For Each TdLin In Itr(TdLy)
    Sk = SkFny(TdLin)
    If Sz(Sk) > 0 Then
        PushI SqyCrtSk, SqlCrtSk(T1(TdLin), Sk)
    End If
Next
End Function

Private Function SkFny(TdLin) As String()
Dim P%, T$, Rst$
P = InStr(TdLin, "|")
If P = 0 Then Exit Function
AsgTRst TakBef(TdLin, "|"), T, Rst
Rst = Replace(Rst, T, "*")
SkFny = SySsl(Rst)
End Function

Private Function TdAy(TdLy$(), A As EF) As Dao.TableDef()
Dim T
For Each T In T1Ay(TdLy)
    PushObj TdAy, Td(T, A)
Next
End Function
Private Function HasPk(TdLin) As Boolean
HasPk = HasSubStr(TdLin, " *Id ")
End Function

Private Function Td(TdLin, A As EF) As Dao.TableDef
'FDic is FDicSmf
Dim O As Dao.TableDef
Dim IsPk As Boolean
Dim F
For Each F In FnyzTdLin(TdLin)
    If HasPk(TdLin) Then
        O.Fields.Append FdzPk(F)
    Else
        O.Fields.Append FdzEF(F, A)
    End If
Next
O.Name = T1(TdLin)
End Function
Function T1z_Itm_T1LikssAy$(Itm, T1LikssAy$())
Dim L, Likss$, T1$
For Each L In T1LikssAy
    AsgTRst L, T1, Likss
    If HitLikss(Itm, Likss) Then T1z_Itm_T1LikssAy = T1: Exit Function
Next
End Function

Private Function FdStr$(F, A As EF)
Dim E$, EStr$
E = T1z_Itm_T1LikssAy(F, A.FldLy)
EStr = FstEleT1(A.EleLy, E)
FdStr = F & " " & EStr
End Function

Private Function FdzEF(F, A As EF) As Dao.Field2
Set FdzEF = FdzFdStr(FdStr(F, A))
End Function

Private Sub Z_CrtSchm()
Dim Schm$(), Db As Database
Set Db = TmpDb
Schm = SampSchm
GoSub Tst
Exit Sub
Tst:
    CrtSchm Db, Schm
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
Dim Td As Dao.TableDef
Dim B As Dao.Field2
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
    Dim FdAy() As Dao.Field
    Set B = Fd("B", dbInteger)
    PushObj FdAy, FdzId("#Tmp")
    PushObj FdAy, Fd("A", dbInteger)
    PushObj FdAy, B
    Set Td = NewTd("#Tmp", FdAy, "A B")
    Return
End Sub

Sub AA()
Z
End Sub
Private Sub Z()
Z_CrtSchm
End Sub
