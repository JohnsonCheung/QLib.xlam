Attribute VB_Name = "MApp_Fun"
Option Explicit
Private X_Acs As New Access.Application
Public Type StrConst
    Nm As String
    Str As String
End Type
Public Type ConstVal
    IsPrv As Boolean
    Nm As String
    TyChr As String
    Ty As String
    Str As String
    NonStrVal As String
End Type
Public Type ConstValOpt
    Som As Boolean
    V As ConstVal
End Type
Public Type StrConsts:    N As Long: Ay() As StrConst:    End Type
Public Type ConstVals:    N As Long: Ay() As ConstVal:    End Type
Function AppDb(Apn) As Database
Set AppDb = Db(AppFb(Apn))
End Function

Function OupFxzNxt$(Apn)
OupFxzNxt = NxtFfn(OupFx(Apn))
End Function

Function OupFx1$()
'Dim A$, B$
'A = OupPth & FmtQQ("TaxExpCmp ?.xlsx", Format(Now, "YYYY-MM-DD HHMM"))
'OupFx1 = A
End Function

Function OupFx$(Apn)
OupFx = PnmOupPth(AppDb(Apn)) & Apn & ".xlsx"
End Function

Function AppFb$(Apn)
AppFb = AppHom & Apn & ".app.accdb"
End Function

Property Get AppHom$()
Static Y$
If Y = "" Then
    Y = AddFdrEns(AddFdrEns(ParPth(TmpRoot), "Apps"), "Apps")
End If
AppHom = Y
End Property
Sub Ens()
'EnsMdCSub
'EnsMdOptExp
'EnsMdSubZZZ
Srt
End Sub

Property Get AutoExec()
'D "AutoExec:"
'D "-Before LnkCcm: CnSy--------------------------"
'D CnSy
'D "-Before LnkCcm: Srcy--------------------------"
'D Srcy
'
'EnsTblSpec

LnkCcm CurrentDb, IsDev
'D "-After LnkCcm: CnSy--------------------------"
'D CnSy
'D "-After LnkCcm: Srcy--------------------------"
'D Srcy
End Property
Function StrConstVal$(Lin)
Dim L$: L = Lin
ShfMthMdy L
If Not ShfPfx(L, "Const") Or (ShfNm(L) = "") Then Thw CSub, "Lin is not a StrConstLin", "Lin", Lin
If Not ShfPfx(L, "$") Then Thw CSub, "No $ after name", "Lin", Lin
If Not ShfPfx(L, " = """) Then Thw CSub, "No [ = ""] after [$]", "Lin", Lin
Dim P&: P = InStr(L, """"): If P = 0 Then Thw CSub, "Should  have 2 dbl-quote", "Lin", Lin
StrConstVal = Left(L, P - 1)
End Function

Function ConstNm$(Lin)
Dim L$: L = Lin
ShfMthMdy L
If ShfPfx(L, "Const") Then ConstNm = TakNm(L)
End Function
Function DocWsInPj() As Worksheet
Set DocWsInPj = DocWszPj(CurPj)
End Function
Function DocWszPj(A As VBProject) As Worksheet
Dim O As New Worksheet
Set O = NewWs("TermDoc")
RgzSq DocSqzPj(A), A1zWs(O)
FmtDocWs O
Set DocWszPj = O
End Function
Private Sub FmtDocWs(DocWs As Worksheet)

End Sub
Private Function DocSqzPj(A As VBProject) As Variant()
Dim Dry(), K, D As Dictionary, TermAset As Aset
PushI Dry, SySsl("Term Lnk DfnStmt")
Set D = DocDiczPj(A)
Set TermAset = AsetzItr(D.Keys)
For Each K In D.Keys
    PushIAy Dry, DryzDocDfn(K, D(K), TermAset)
Next
DocSqzPj = SqzDry(Dry)
End Function

Function DryzDocDfn(Nm, Lines, TermAset As Aset) As Variant()  'DocDfn = Nm + Lines.  DocDry = Term + Lnk + Stmt
Dim S$(): S = StmtLy(Lines)
Dim N0$(): N0 = AywDist(NyzStr(Lines))
Dim N1$(): N1 = AywInAset(N0, TermAset)
Dim N$(): N = AyeEle(N1, Nm)
Dim J%, Term, Nm1, Stmt, UN%, US%
UN = UB(N)
US = UB(S)
Dim O()
ReDim O(1 To Max(US, UN) + 1, 1 To 3)
O(1, 1) = Nm
For J = 0 To US
    O(J + 1, 3) = S(J)
Next
For J = 0 To UN
    O(J + 1, 2) = N(J)
Next
DryzDocDfn = DryzSq(O)
End Function

Function DocDicInPj() As Dictionary
Static X As Dictionary
If IsNothing(X) Then Set X = DocDiczPj(CurPj)
Set DocDicInPj = X
End Function

Function StrConstszConstVals(A As ConstVals) As StrConsts
Dim O As StrConsts, J%
For J = 0 To A.N - 1
    With A.Ay(J)
    If .Str <> "" Then
        PushStrConst O, StrConst(.Nm, .Str)
    End If
    End With
Next
End Function
Function StrConst(Nm$, Str$) As StrConst
StrConst.Nm = Nm
StrConst.Str = Str
End Function
Sub PushStrConst(O As StrConsts, M As StrConst)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub
Function ConstVals(Ly$()) As ConstVals
Dim L, O As ConstVals
For Each L In Itr(ContLyzLy(Ly))
    ConstValsPushOpt O, ConstValOpt(L)
Next
ConstVals = O
End Function
Sub ConstValsPushOpt(O As ConstVals, M As ConstValOpt)
If Not M.Som Then Exit Sub
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M.V
O.N = O.N + 1
End Sub
Function ConstValOpt(Lin) As ConstValOpt
If HasSfx(Lin, "_") Then Thw CSub, "Cannot has Sfx [_]", "Lin", Lin
Dim L$, N$: L = Lin
Dim O As ConstValOpt
O.V.IsPrv = ShfMthMdy(L) = "Private"
If Not ShfPfx(L, "Const ") Then Exit Function
O.V.Nm = ShfNm(L)
O.V.TyChr = ShfTyChr(L)
If ShfTermX(L, "As") Then
    O.V.Ty = ShfT1(L)
    O.V.NonStrVal = L
    Exit Function
End If
If Not ShfTermX(L, "=") Then Thw CSub, "Lin is invalid const line: no [ = ] after name", "Lin", Lin
If ShfPfx(L, """") Then
    Dim P&: P = InStr(L, """"): If P = 0 Then Thw CSub, "Something wrong in Lin, which is supposed to be string const lin.  There is no snd [""]", "Lin", Lin
    O.V.Str = Left(L, P - 1)
Else
    O.V.NonStrVal = L
End If
O.Som = True
ConstValOpt = O
End Function
Function IsStrConst(A As ConstValOpt) As Boolean
With A
    If .Som Then
        If .V.Str <> "" Then
            IsStrConst = True
        End If
    End If
End With
End Function

Function StrConsts(Ly$()) As StrConsts
Dim A As ConstVals, M As ConstVal, J&, O As StrConsts
A = ConstVals(Ly)
For J = 0 To A.N - 1
    M = A.Ay(J)
    If M.Str <> "" Then
        PushStrConst O, StrConst(M.Nm, M.Str)
    End If
Next
StrConsts = O
End Function

Function DocDiczDcl(Dcl$) As Dictionary
Dim DocNm$, O As New Dictionary
Set O = New Dictionary
Dim A As StrConsts, J&, M As StrConst
A = StrConsts(SplitCrLf(Dcl))
For J = 0 To A.N - 1
    M = A.Ay(J)
    If HasPfx(M.Nm, "DocOf") Then
        DocNm = Mid(M.Nm, 6)
        O.Add DocNm, M.Str
    End If
Next
Set DocDiczDcl = O
End Function


Property Get IsDev() As Boolean
Static X As Boolean, Y As Boolean
If Not X Then
    X = True
    Y = Not HasPth(ProdPth)
End If
IsDev = Y
End Property

Property Get IsProd() As Boolean
IsProd = Not IsDev
End Property


Function PgmDb_DtaDb(A As Database) As Database
Set PgmDb_DtaDb = DBEngine.OpenDatabase(PgmDb_DtaFb(A))
End Function

Function PgmDb_DtaFb$(A As Database)
End Function

Property Get ProdPth$()
ProdPth = "N:\SAPAccessReports\"
End Property

Private Sub ZZ()
Dim A As Database
PgmDb_DtaDb A
PgmDb_DtaFb A
End Sub

Property Let ApnzDb(A As Database, V$)
ValOfQ(A, SqlSel_F("Apn")) = V
End Property

Property Get ApnzDb$(A As Database)
ApnzDb = ValOfQ(A, "Select Apn from Apn")
End Property

