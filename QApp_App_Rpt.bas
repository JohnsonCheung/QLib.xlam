Attribute VB_Name = "QApp_App_Rpt"
Option Compare Text
Option Explicit
Private Const CMod$ = "BRpt."
Private Type CpyInpPm
    IsCpyInp As Boolean  ' Is cpy inp ws to oup wb
    OupWb As Workbook
    WrkDb As Database
    LnkImpSrc() As String
End Type
Type RptPm
    Appn As String
    Appv As String
    InpFilSrc() As String
    LnkImpSrc() As String
    IsCpyInp As Boolean
    OupFx As String
    WbFmtr As IWbFmtr
    OupGenr As IOupGenr
End Type

Sub Rpt(A As RptPm) 'Gen&Vis OupFx using LidPm as NxtFfnzAva.
Dim Init:                                   SetApp A.Appn, A.Appv
Dim WrkPth$:                       WrkPth = WPth
Dim O1:                                     CpyFfnyzIfDif A.InpFilSrc, WrkPth          ' <== Cpy inp fil to wpth
Dim WrkFb$:                         WrkFb = WFb
Dim O2a:                                    DltFfnIf WrkFb
Dim O2b:                                    CrtFb WrkFb                                ' <== Crt wrk fb
Dim WrkDb As Database:          Set WrkDb = WrkDb
Dim O3:                                     LnkImp Sy(A.InpFilSrc, A.LnkImpSrc), WrkDb ' <== LnkImp
Dim O4:                                     A.OupGenr.GenOupTblFmTmpInp WrkDb             ' <== Gen oup tbl
Dim OupFx$:                         OupFx = AppOupFx
Dim O5:                                     ExpAppTp
Dim O6:                                     RfhFx OupFx, WrkFb                         ' <== Crt oup fx
Dim OupWb As Workbook:          Set OupWb = WbzFx(OupFx)
Dim X     As CpyInpPm:             X.IsCpyInp = A.IsCpyInp
                                            Set X.OupWb = OupWb
                                            Set X.WrkDb = WrkDb
                                            X.LnkImpSrc = A.LnkImpSrc
Dim O7:                                     XCpyInp X                         ' <== Optional cpy inp tbl to oup wb
Dim O8:                                     A.WbFmtr.FmtWb OupWb                       ' <== Fmt oup wb
Dim O9a:                                    OupWb.Save                                 ' <== Sav
Dim O9b:                                    ShwWb OupWb                                ' <== Set Vis
End Sub


Private Sub XCpyInp(A As CpyInpPm)
Dim J%

End Sub



Private Sub Z_Rpt()
Dim WDb As Database, B As IOupGenr
GoSub ZZ
Exit Sub
ZZ:
    Set WDb = Nothing
    Set B = New OupGenrzTaxCmp
    GoTo Tst
Tst:
    'Rpt_Oup WDb, B
    Return
End Sub
