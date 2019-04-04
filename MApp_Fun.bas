Attribute VB_Name = "MApp_Fun"
Option Explicit
Private X_Acs As New Access.Application
Function AppDb(Apn) As Database
Set AppDb = Db(AppFb(Apn))
End Function

Function OupFxzNxt$(Apn)
OupFxzNxt = NxtFfn(OupFx(Apn))
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
Function DocLy(DclLy$()) As String()
Dim Lin, N$
For Each Lin In Itr(DclLy)
    N = ConstNm(Lin)
    If Left(N, 5) = "DocOf" Then PushI DocLy, Mid(N, 6) & " " & StrConstVal(Lin)
Next
End Function
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
Function DocDicOfPj() As Dictionary
Static X As Dictionary
If IsNothing(X) Then Set X = DocDiczPj(CurPj)
Set DocDicOfPj = X
End Function
Function IsDocNm(S) As Boolean
If Not IsNm(S) Then Exit Function
IsDocNm = Left(S, 5) = "DocOf"
End Function
Sub AsgStrConstNmAaaVal(OStrConstNm$, OStrConstVal$, Lin)
Dim L$, N$: L = Lin
OStrConstNm = ""
OStrConstVal = ""
ShfMthMdy L
If Not ShfPfx(L, "Const ") Then Exit Sub
N = ShfNm(L)
If Not ShfPfx(L, "$") Then Exit Sub
OStrConstNm = N
If Not ShfPfx(L, " = """) Then Thw CSub, "Something wrong in Lin, which is suppose to be a string const lin.  There is no [ = ""] after [$]", "Lin", Lin
Dim P&: P = InStr(L, """"): If P = 0 Then Thw CSub, "Something wrong in Lin, which is supposed to be string const lin.  There is no snd [""]", "Lin", Lin
OStrConstVal = Left(L, P - 1)
End Sub

Function DocDiczDcl(Dcl) As Dictionary
Dim MdDNm, Lin, DocNm$, StrConstVal$, StrConstNm$
Set DocDiczDcl = New Dictionary
For Each Lin In LinItr(Dcl)
    AsgStrConstNmAaaVal StrConstNm, StrConstVal, Lin
    If Left(StrConstNm, 5) = "DocOf" Then
        DocNm = Mid(StrConstNm, 6)
        If DocDiczDcl.Exists(DocNm) Then Thw CSub, "DocNm is dup", "DocNm", DocNm
        DocDiczDcl.Add DocNm, StrConstVal
    End If
Next
End Function

Function DocDiczPj(A As VBProject) As Dictionary
Dim O As New Dictionary, Dcl
For Each Dcl In DclDiczPj(A).Items
    PushDic O, DocDiczDcl(Dcl)
Next
Set DocDiczPj = O
End Function
Sub Doc(Nm$)
If DocDicOfPj.Exists(Nm) Then D DocDicOfPj(Nm) Else D "Not exist"
'#BNmMIS is Method-B-Nm-of-Missing.
'           Missing means the method is found in FmPj, but not ToPj
'#FmDicB is MthDic-of-MthBNm-zz-MthLines.   It comes from FmPj
'#ToDicA is MthDic-of-MthANm-zz-MthLinesAy. It comes from ToPj
'#ToDicAB is ToDicA and FmDicB
'#ANm is method-a-name, NNN or NNN:YYY
'        If the method is Sub|Fun, just MthNm
'        If the method is Prp    ,      MthNm:MthTy
'        It is from ToPj (#ToA)
'        One MthANm will have one or more MthLines
'#BNm is method-b-name, MMM.NNN or MMM.NNN:YYY
'        MdNm.MthNm[:MthTy]
'        It is from FmPj (#BFm)
'        One MthBNm will have only one MthLines
'#Missing is for each MthBNm found in FmDicB, but its MthNm is not found in any-method-name-in-ToDicA
'#Dif is for each MthBNm found in FmDicB and also its MthANm is found in ToDicA
'        and the MthB's MthLines is dif and any of the MthA's MthLines
'       (Note.MthANm will have one or more MthLines (due to in differmodule))
End Sub

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

