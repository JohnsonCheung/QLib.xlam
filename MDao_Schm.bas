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
Private Type FdRslt
    Som As Boolean
    Fd As DAO.Field2
End Type

Sub CrtSchmzVbl(A As Database, SchmVbl$)
CrtSchm A, SplitVBar(SchmVbl)
End Sub

Sub CrtSchm(A As Database, Schm$())
Const CSub$ = CMod & "CrtSchm"
ThwIfErMsg ErzSchm(Schm), CSub, "there is error in the Schm", "Schm Db", SyAddIxPfx(Schm, 1), DbNm(A)
Dim TdLy$():            TdLy = AywRmvT1(Schm, C_Tbl)
Dim EF As EF:             EF = EFzSchm(Schm)
Dim T() As DAO.TableDef:   T = TdAy(TdLy, EF)
Dim P$():                  P = SqyCrtPkzTny(PkTny(TdLy))
Dim S$():                  S = SqyCrtSk(TdLy)
Dim DicT As Dictionary: Set DicT = Dic(AywRmvTT(Schm, C_Des, C_Tbl))
Dim DicF As Dictionary: Set DicF = Dic(AywRmvTT(Schm, C_Des, C_Fld))
                   AppTdAy A, T
                   RunSqy A, P
                   RunSqy A, S
Set TblDesDic(A) = DicT
Set FldDesDic(A) = DicF
End Sub

Private Function EFzSchm(Schm$()) As EF
EFzSchm.EleLy = AywRmvT1(Schm, "Ele")
EFzSchm.FldLy = AywRmvT1(Schm, "Fld")
End Function

Private Function PkTny(TdLy$()) As String()
Dim I, L$
For Each I In TdLy
    L = I
    If HasSubStr(L, " *Id ") Then
        PushI PkTny, T1(L)
    End If
Next
End Function

Private Function SqyCrtSk(TdLy$()) As String()
Dim TdLin$, I, Sk$()
For Each I In Itr(TdLy)
    TdLin = I
    Sk = SkFny(TdLin)
    If Si(Sk) > 0 Then
        PushI SqyCrtSk, SqlCrtSk_T_SkFny(T1(TdLin), SyRplStar(Sk, T1(TdLin)))
    End If
Next
End Function

Private Function SkFny(TdLin$) As String()
Dim P%, T$, Rst$
P = InStr(TdLin, "|")
If P = 0 Then Exit Function
AsgTRst Bef(TdLin, "|"), T, Rst
Rst = Replace(Rst, T, "*")
SkFny = SySsl(Rst)
End Function

Private Function TdAy(TdLy$(), A As EF) As DAO.TableDef()
Dim TdLin$, I
For Each I In TdLy
    TdLin = I
    PushObj TdAy, TdzLin(TdLin, A)
Next
End Function

Private Function TdzLin(TdLin$, A As EF) As DAO.TableDef
Dim T: T = T1(TdLin)
Dim O As DAO.TableDef: Set O = New DAO.TableDef
O.Name = T
Dim F$, I, Fd As DAO.Field2
For Each I In FnyzTdLin(TdLin)
    F = I
    If F = T & "Id" Then
        Set Fd = FdzPk(F)
    Else
        Set Fd = FdzEF(F, A)
    End If
    O.Fields.Append Fd
Next
Set TdzLin = O
End Function

Function T1zLookup_FmT1LikssAy_ByItm$(Itm$, T1LikssAy$())
Dim L, Likss$, T1$
For Each L In T1LikssAy
    AsgTRst L, T1, Likss
    If HitLikss(Itm, Likss) Then T1zLookup_FmT1LikssAy_ByItm = T1: Exit Function
Next
End Function

Private Function FdzEF(F$, A As EF) As DAO.Field2
If Left(F, 2) = "Id" Then Stop
Dim Ele$: Ele = T1zLookup_FmT1LikssAy_ByItm(F, A.FldLy)
If Ele <> "" Then Set FdzEF = FdzEle(Ele, A.EleLy, F): Exit Function
Set FdzEF = FdzFld(F):                    If Not IsNothing(FdzEF) Then Exit Function
Set FdzEF = FdzEle(CStr(F), A.EleLy, F):  If Not IsNothing(FdzEF) Then Exit Function
Thw CSub, FmtQQ("Fld(?) not in EF and not StdFld", F)
End Function

Private Function FdzEle(Ele$, EleLy$(), F$) As DAO.Field2
Dim EStr$: EStr = EleStr(EleLy, Ele)
If EStr <> "" Then Set FdzEle = FdzFdStr(F & " " & EStr): Exit Function
Set FdzEle = FdzShtTys(Ele, F): If Not IsNothing(FdzEle) Then Exit Function
EStr = EleStr(EleLy, F)
If EStr <> "" Then Set FdzEle = FdzFdStr(F & " " & EStr): Exit Function
Set FdzEle = FdzShtTys(F, F)
Dim EleNy$(): EleNy = T1Sy(EleLy)
Thw CSub, FmtQQ("Fld(?) of Ele(?) not found in EleLy-of-EleAy(?) and not StdEle", F, Ele, TLin(EleNy))
End Function

Private Function EleStr$(EleLy$(), Ele$)
EleStr = RmvT1(FstElezT1(EleLy, Ele))
End Function

Private Function EleStrzStd$(Ele)
End Function

Private Property Get Schm1() As String()
Erase XX
X "Tbl A *Id *Nm     | *Dte AATy Loc Expr Rmk"
X "Tbl B *Id  AId *Nm | *Dte"
X "Fld Txt AATy"
X "Fld Mem Rmk"
X "Ele Loc Txt Rq Dft=ABC [VTxt=Loc must cannot be blank] [VRul=IsNull([Loc]) or Trim(Loc)='']"
X "Ele Expr Txt [Expr=Loc & 'abc']"
X "Des Tbl  A     AA BB "
X "Des Tbl  A     CC DD "
X "Des Fld  ANm   AA BB "
X "Des Fld  A.ANm TF_Des-AA-BB"
Schm1 = XX
Erase XX
End Property

Private Sub Z_CrtSchm()
Dim D As Database, Schm$()
GoSub T1
Exit Sub

T1:
    Set D = TmpDb
    Schm = Schm1
    GoTo Tst
Tst:
    CrtSchm D, Schm
    Dmp TdLyzDb(D)
    Return
End Sub

Sub AA()
Z
End Sub

Private Sub Z()
Z_CrtSchm
End Sub

Sub AppTdAy(A As Database, TdAy() As DAO.TableDef)
Dim T
For Each T In Itr(TdAy)
    A.TableDefs.Append T
Next
End Sub


Function FnyzTdLin(TdLin) As String()
Dim T$, Rst$
AsgTRst TdLin, T, Rst
If HasSfx(T, "*") Then
    T = RmvSfx(T, "*")
    Rst = T & "Id " & Rst
End If
Rst = Replace(Rst, "*", T)
Rst = Replace(Rst, "|", " ")
FnyzTdLin = SySsl(Rst)
End Function

