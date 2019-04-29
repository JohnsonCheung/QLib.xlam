Attribute VB_Name = "MApp_Rpt"
Option Explicit
Type CpyOneWsPm: InpWsn As String: AsWsn As String: InpColNy() As String: AsFny() As String: End Type
Type CpyOneWbPm: InpFx As String: N As Integer: Ay() As CpyOneWsPm: End Type
Type CpyInpWsPm: N As Integer: Ay() As CpyOneWbPm: End Type
Type RptPm
    Apn As String
    CpyInpWsPm As CpyInpWsPm
    LnkImpPm As LnkImpPm
    OupFx As String
    WbFmtr As IWbFmtr
    Genr As IGenr
End Type

Sub Rpt(A As RptPm) 'Gen&Vis OupFx using LidPm as NxtFfn.
With A
    Rpt_Cpy .InpFxAy, .WPth
    Rpt_Ini .WFb
    Dim W As Database: Set W = Db(.WFb)
    LnkImp A.LnkPm
    Rpt_Oup W, .Genr
    Rpt_Gen W, .OupFx
    Dim OupWb As Workbook: Set OupWb = WbzFx(.OupFx)
    Rpt_CpyWs OupWb, .CpyInpWsPm
    Rpt_Fmt OupWb, .WbFmtr
End With
OupWb.Save
VisWb OupWb
End Sub
Private Sub Rpt_Ini(WFb$)
DltFfnIf WFb
CrtFb WFb
End Sub
Private Sub Rpt_Cpy(InpFxAy$(), ToPth$)
CpyFilzIfDif InpFxAy, ToPth
End Sub
Private Sub Rpt_Fmt(OupWb As Workbook, A As IWbFmtr)
A.FmtWb OupWb
End Sub

Private Sub Rpt_LnkFx(W As Database, InpFxAy$(), InpFxTny$(), InpWsNy$())
Dim J%
For J = 0 To UB(InpFxAy)
    LnkFxw W, InpFxTny(J), InpFxAy(J), InpWsNy(J)
Next
End Sub
Private Sub Rpt_CpyWs(OupWb As Workbook, B As CpyInpWsPm)
Dim J%
For J = 0 To B.N - 1
    CpyWs_CpyOneWb OupWb, B.Ay(J)
Next
End Sub
Private Sub CpyWs_CpyOneWb(OupWb As Workbook, A As CpyOneWbPm)
With A
    Dim J%
    Dim FmWb As Workbook: Set FmWb = WbzFx(.InpFx)
    For J = 0 To A.N - 1
        CpyOneWb_CpyOneWs FmWb, OupWb, A.Ay(J)
    Next
    FmWb.Close False
End With
End Sub
Private Sub CpyOneWb_CpyOneWs(FmWb As Workbook, OupWb As Workbook, A As CpyOneWsPm)
With A
    WszWb(FmWb, .InpWsn).Copy , LasWs(OupWb)
    SetWsNm LasWs(OupWb), .AsWsn
    Dim Ws As Worksheet: Set Ws = WszWb(OupWb, .AsWsn)
    CpyOneWs_PasteAsVal Ws
    CpyOneWs_RmvCol Ws, .InpColNy
    CpyOneWs_InsFnyRow Ws, .AsFny
    CpyOneWs_SetLo Ws
End With
End Sub
Private Sub CpyOneWs_SetLo(Ws As Worksheet)

End Sub
Private Sub CpyOneWs_InsFnyRow(Ws As Worksheet, AsFny$())
CvRg(Ws.Rows(2)).EntireRow.Insert
RgzAyH AsFny, WsRC(Ws, 2, 1)
End Sub
Function FnyzWs(A As Worksheet)

End Function
Private Sub CpyOneWs_RmvCol(Ws As Worksheet, InpColNy$())
Dim Fny$(): Fny = ReserveAy(FnyzWs(Ws))
Dim Col%: Col = Si(Fny)
Dim F
For Each F In Itr(Fny)
    If Not HasEle(InpColNy, F) Then
        WsC(Ws, Col).EntireColumn.Delete
    End If
    Col = Col - 1
Next
End Sub

Private Sub CpyOneWs_PasteAsVal(A As Worksheet)

End Sub
Private Sub PasteWsAsVal(Ws As Worksheet)

End Sub

Private Sub Rpt_Gen(W As Database, OupFx$)

End Sub
Private Sub Rpt_LnkFb(W As Database, InpFbAy$(), InpFbTny$())

End Sub
Private Sub Rpt_ImpFx(W As Database, ImpWsSqy$())
RunSqy W, ImpWsSqy
End Sub
Private Sub Rpt_ImpFb(W As Database, InpFbTny$())
RunSqy W, ImpFb_Sqy(InpFbTny)
End Sub
Private Function ImpFb_Sqy(InpFbTny$()) As String()

End Function

Sub Z_Rpt_Oup()
Dim WDb As Database, B As IGenr
GoSub ZZ
Exit Sub
ZZ:
    Set WDb = Nothing
    Set B = New ATaxExpCmp_OupTblGenr
    GoTo Tst
Tst:
    Rpt_Oup WDb, B
    Return
End Sub
Private Sub Rpt_Oup(WDb As Database, B As IGenr)
B.GenOupTblFmTmpInp WDb
End Sub
