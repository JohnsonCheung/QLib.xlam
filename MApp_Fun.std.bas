Attribute VB_Name = "MApp_Fun"
Option Explicit
Const CMod$ = "MApp___Fun."
Private X_Acs As New Access.Application
Function AppDb(Apn) As Database
Set AppDb = Db(AppFb(Apn))
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

LnkCcmz CDb, IsDev
'D "-After LnkCcm: CnSy--------------------------"
'D CnSy
'D "-After LnkCcm: Srcy--------------------------"
'D Srcy
End Property

Sub Doc()
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
Doc
Ens
PgmDb_DtaDb A
PgmDb_DtaFb A
End Sub

Property Let ApnzDb(A As Database, V$)
ValzDbq(A, SqlSel_F("Apn")) = V
End Property

Property Get ApnzDb$(A As Database)
ApnzDb = ValzDbq(A, "Select Apn from Apn")
End Property

Property Get AppFbAy() As String()
Push AppFbAy, AppJJFb
Push AppFbAy, AppStkShpCstFb
Push AppFbAy, AppStkShpRateFb
Push AppFbAy, AppTaxExpCmpFb
Push AppFbAy, AppTaxRateAlertFb
End Property
