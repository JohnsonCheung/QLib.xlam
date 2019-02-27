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
    Fd As Dao.Field2
End Type
Sub CrtSchm(A As Database, Schm$())
Const CSub$ = CMod & "CrtSchm"
'ThwErMsg ErzSchm(Schm), CSub, "there is error in the Schm", "Schm Db", AyAddIxPfx(Schm, 1), DbNm(A)
Dim TdLy$():          TdLy = AywRmvT1(Schm, C_Tbl)
Dim EF As EF:           EF = EFzSchm(Schm)
Dim T() As Dao.TableDef: T = TdAy(TdLy, EF)
Dim P$():                P = SqyCrtPkzTny(PkTny(TdLy))
Dim S$():                S = SqyCrtSk(TdLy)
Dim DT As Dictionary: Set DT = Dic(AywRmvTT(Schm, C_Des, C_Tbl))
Dim DF As Dictionary: Set DF = Dic(AywRmvTT(Schm, C_Des, C_Fld))
Stop
                      AppTdAy A, T
                      RunSqy A, P
                      RunSqy A, S
   Set TblDesDic(A) = DT
   Set FldDesDic(A) = DF
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
Dim TdLin
For Each TdLin In TdLy
    PushObj TdAy, Td(TdLin, A)
Next
End Function

Private Function Td(TdLin, A As EF) As Dao.TableDef
'FDic is FDicSmf
Dim O As New Dao.TableDef: O.Name = T1(TdLin)
Dim T: T = O.Name
Dim F
For Each F In FnyzTdLin(TdLin)
    O.Fields.Append FdzEF(F, T, A)
Next
Set Td = O
End Function

Function T1z_Itm_T1LikssAy$(Itm, T1LikssAy$())
Dim L, Likss$, T1$
For Each L In T1LikssAy
    AsgTRst L, T1, Likss
    If HitLikss(Itm, Likss) Then T1z_Itm_T1LikssAy = T1: Exit Function
Next
End Function

Private Function FdzEF(F, T, A As EF) As Dao.Field2
Dim Ele$: Ele = T1z_Itm_T1LikssAy(F, A.FldLy)
If Ele <> "" Then Set FdzEF = FdzEle(Ele, A.EleLy, F): Exit Function
Set FdzEF = FdzStdFld(F)
If Not IsNothing(FdzEF) Then Exit Function
Thw CSub, FmtQQ("Fld(?) not in EF and not StdFld", F)
End Function

Private Function FdzEle(Ele$, EleLy$(), F) As Dao.Field2
Dim EStr$: EStr = EleStr(EleLy, Ele)
If EStr <> "" Then Set FdzEle = FdzFdStr(F & " " & EStr): Exit Function
Set FdzEle = FdzStdEle(Ele, F): If Not IsNothing(FdzEle) Then Exit Function
EStr = FdzStdEle(F, F)
If EStr <> "" Then Set FdzEle = FdzFdStr(F & " " & EStr): Exit Function
Dim EleNy$(): EleNy = T1Ay(EleLy)
Thw CSub, FmtQQ("Fld(?) of Ele(?) not found in EleLy-of-EleAy(?) and not StdEle", F, Ele, TLin(EleNy))
End Function

Private Function EleStr$(EleLy$(), Ele$)
EleStr = FstEleT1(EleLy, Ele)
End Function

Private Function EleStrzStd$(Ele)

End Function

Private Function FdzStdFldNm(F) As Dao.Field2

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
    Schm = Schm1
    GoTo Tst
Tst:
    CrtSchm D, Schm
    Return
End Sub

Sub AA()
Z
End Sub
Private Sub Z()
Z_CrtSchm
End Sub

Private Sub AppTdAy(A As Database, TdAy() As Dao.TableDef)
Dim T
For Each T In Itr(TdAy)
    A.TableDefs.Append T
Next
End Sub

