VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Const CLib$ = "QApp."
Const CMod$ = CLib & "App."
Private Type A
    Appn As String
    Appv As String
    Fb As String
    Db As Database
End Type
Private A As A
Property Get Appn$()
Appn = A.Appn
End Property
Property Get Appv$()
Appv = A.Appv
End Property
Sub SetApp(Appn$, Appv$)
A.Appn = Appn
A.Appv = Appv
'A.Fb = WFb
DltFfnIf Fb
CrtFb Fb
Set A.Db = Db(A.Fb)
End Sub

Property Get Pth$()
Pth = MxFfn.Pth(CPjf)
End Property

Function TpFx$()
Stop
'TpFx = AppPth & A.Appn & "(Template).xlsx"
End Function

Function TpFxm$()
Stop
'AppTpFxm = AppPth & A.Appn & "(Template).xlsm"
End Function

Sub OpnTp()
OpnFx Tp
End Sub

Function Tp$()
Dim F$
F = TpFxm: If HasFfn(F) Then Tp = F: Exit Function
F = TpFx:  If HasFfn(F) Then Tp = F: Exit Function
End Function

Function OupFxzNxt$()
OupFxzNxt = NxtFfnzAva(OupFx)
End Function

Function OupPth$()
OupPth = VzPm(A.Db, "OupPth")
End Function

Function FfnzPm$(PmNm$)
FfnzPm = VzPm(A.Db, PmNm & "Ffn")
End Function

Function OupFx$()
OupFx = OupPth & A.Appn & ".xlsx"
End Function

Function Fb$()
'C:\Users\user\Desktop\MHD\SAPAccessReports\TaxExpCmp\TaxExpCmp\TaxExpCmp.1_3..accdb
Fb = Hom & JnDotAp(Appn, Appv, "app", "accdb")
End Function

Function Hom$()
Static Y$
If Y = "" Then Y = AddFdrEns(AddFdrEns(ParPth(TmpRoot), "Apps"), "Apps")
Hom = Y
End Function

Function AutoExec()
'D "AutoExec:"
'D "-Before LnkCcm: CnSy--------------------------"
'D CnSy
'D "-Before LnkCcm: Srcy--------------------------"
'D Srcy
'
'EnsTblSpec

LnkCcm CurrentDb, CUsr = "User"
'D "-After LnkCcm: CnSy--------------------------"
'D CnSy
'D "-After LnkCcm: Srcy--------------------------"
'D Srcy
End Function

Sub ImpTp()
Dim T$: T = Tp
Const CSub$ = CMod & "ImpTp"
If T = "" Then
    Inf CSub, "Tp not exist WPth, no Import", "AppNm Tp WPth", A.Db, Appn, T, Pth
    Exit Sub
End If
Dim D As Database: Set D = A.Db
If IsAttOld(D, "Tp", T) Then ImpAtt D, "Tp", T '<== Import
End Sub

Function TpWsCdNy() As String()
'TpWsCdNy = WszFxwCdNy(TpFx)
End Function

'==============================================
Function TpPth$()
TpPth = EnsPth(TmpHom & "Template\")
End Function

Sub RfhTpWc()
RfhFx Tp, Fb
End Sub

Function TpWb() As Workbook
Set TpWb = WbzFx(Tp)
End Function

Function WcsyzTp() As String()
Dim W As Workbook, X As Excel.Application
Set X = New Excel.Application
Set W = X.Workbooks.Open(Tp)
'TpWcsy = WcStrAyWbOLE(W)
W.Close False
Set W = Nothing
X.Quit
Set X = Nothing
End Function
Sub ExpTp()
ExpAtt A.Db, "Tp", Tp
End Sub

Function PthzPm$(D As Database, PmNm$)
PthzPm = EnsPthSfx(VzPm(D, PmNm & "Pth"))
End Function

Function Pjfnm$(D As Database, PmNm$)
Pjfnm = VzPm(D, PmNm & "Fn")
End Function

Function VzPm(D As Database, PmNm$)
Dim Q$: Q = FmtQQ("Select ? From Pm where CUsr='?'", PmNm, CUsr)
VzPm = FvzQ(D, Q)
End Function

Sub SetVzPm(D As Database, PmNm$, V)
With D.TableDefs("Pm").OpenRecordset
    .Edit
    .Fields(PmNm).Value = V
    .Update
End With
End Sub

Sub BrwPm(D As Database)
BrwT D, "Pm"
End Sub
