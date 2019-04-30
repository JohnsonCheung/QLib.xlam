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
Const CMod$ = ""
Private Type A
    Nm As String
    Ver As String
    Db As Database
End Type
Private A As A
Friend Function Init(Nm$, Ver$) As App
A.Nm = Nm
A.Ver = Ver
Set Init = Me
If Not HasFfn(Fb$) Then Thw CSub, "App.Fb not exist", "App.Fb", Fb
End Function
Property Get Nm$()
Nm = A.Nm
End Property
Property Get Ver$()
Ver = A.Ver
End Property

Property Get WPth$()
Stop '
End Property
Property Get TpFx$()
TpFx = WPth & A.Nm & "(Template).xlsx"
End Property

Property Get TpFxm$()
TpFxm = WPth & A.Nm & "(Template).xlsm"
End Property
Property Get WFb$()

End Property
Sub OpnTp()
OpnFx Tp
End Sub
Function TpWb() As Workbook
Set TpWb = WbzFx(Tp)
End Function
Property Get Tp$()
Dim A$
A = TpFxm: If HasFfn(A) Then Tp = A: Exit Property
A = TpFx:  If HasFfn(A) Then Tp = A: Exit Property
End Property

Property Get Db() As Database
Set Db = A.Db
End Property

Property Get OupFxzNxt$()
OupFxzNxt = NxtFfn(PmOupFx)
End Property
Property Get OupPth$()
OupPth = ValzPm(Db, "OupPth")
End Property
Property Get FfnzPm$(PmNm$)
FfnzPm = ValzPm(Db, PmNm & "Ffn")
End Property

Property Get PmOupFx$()
PmOupFx = OupPth & A.Nm & ".xlsx"
End Property

Property Get Fb$()
'C:\Users\user\Desktop\MHD\SAPAccessReports\TaxExpCmp\TaxExpCmp\TaxExpCmp.1_3.app.accdb
Fb = Hom & JnDotAp(A.Nm, A.Ver, "app", "accdb")
End Property

Property Get Hom$()
Static Y$
If Y = "" Then Y = AddFdrEns(AddFdrEns(ParPth(TmpRoot), "Apps"), "Apps")
Hom = Y
End Property

Private Sub Class_Terminate()
ClsDb Db
End Sub
