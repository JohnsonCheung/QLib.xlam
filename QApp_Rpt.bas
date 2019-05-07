Attribute VB_Name = "QApp_Rpt"
Option Explicit
Private Const CMod$ = "BRpt."
Type CpyOneWsPm: InpWsn As String: AsWsn As String: InpColNy() As String: AsFny() As String: End Type
Type CpyOneWbPm: InpFx As String: N As Integer: Ay() As CpyOneWsPm: End Type
Type CpyInpWsPm: N As Integer: Ay() As CpyOneWbPm: End Type
Type RptPm
    Apn As String
    InpFilSrc() As String
    LnkImpSrc() As String
    CpyInpWsPm As CpyInpWsPm
    OupFx As String
    WbFmtr As IWbFmtr
    Genr As IGenr
End Type
Private Type B
    WDb As Database
    WFb As String
End Type
Private B  As B
Private A As RptPm
Private Sub R0_B()
'Dim W As Database: Set W = Db(.WFb)
End Sub
Private Sub R(RptPm As RptPm)
A = RptPm
R0_B
R1_CpyInpFx
R2_CrtWFb
R3_LnkImp
R4_GenOupTbl
R5_GenFx
R6_CpyInpWsToOupWs
R7_FmtOupFx
R8_SavAndVis
End Sub
Private Sub R8_SavAndVis()
Dim OupWb As Workbook
OupWb.Save
ShwWb OupWb
End Sub

Sub R3_LnkImp()
LnkImp A.InpFilSrc, A.LnkImpSrc, B.WDb
End Sub
Sub Rpt(A As RptPm) 'Gen&Vis OupFx using LidPm as NxtFfnzAva.
R A
End Sub
Private Sub R2_CrtWFb()
Dim WFb$
DltFfnIf B.WFb
CrtFb B.WFb
End Sub
Private Sub R1_CpyInpFx()
Dim InpFxSy$(), ToPth$
CpyFfnSyzIfDif InpFxSy, ToPth
End Sub
Private Sub R7_FmtOupFx()
Dim OupWb As Workbook, A As IWbFmtr
A.FmtWb OupWb
End Sub

Private Sub R6_CpyInpWsToOupWs()
Dim OupWb As Workbook, B As CpyInpWsPm
Dim J%
For J = 0 To B.N - 1
    R61_CpyOneWb OupWb, B.Ay(J)
Next
End Sub
Private Sub R61_CpyOneWb(OupWb As Workbook, A As CpyOneWbPm)
With A
    Dim J%
    Dim FmWb As Workbook: Set FmWb = WbzFx(.InpFx)
    For J = 0 To A.N - 1
        R611_CpyOneWs FmWb, OupWb, A.Ay(J)
    Next
    FmWb.Close False
End With
End Sub
Private Sub R611_CpyOneWs(FmWb As Workbook, OupWb As Workbook, A As CpyOneWsPm)
With A
    WszWb(FmWb, .InpWsn).Copy , LasWs(OupWb)
    SetWsn LasWs(OupWb), .AsWsn
    Dim Ws As Worksheet: Set Ws = WszWb(OupWb, .AsWsn)
    R6111_PasteAsVal Ws
    R6111_RmvCol Ws, .InpColNy
    R6111_InsFnyRow Ws, .AsFny
    R6111_SetLo Ws
End With
End Sub
Private Sub R6111_SetLo(Ws As Worksheet)

End Sub
Private Sub R6111_InsFnyRow(Ws As Worksheet, AsFny$())
CvRg(Ws.Rows(2)).EntireRow.Insert
RgzAyH AsFny, WsRC(Ws, 2, 1)
End Sub
Function FnyzWs(A As Worksheet) As String()

End Function
Private Sub R6111_RmvCol(Ws As Worksheet, InpColNy$())
Dim Fny$(): Fny = Reverse(FnyzWs(Ws))
Dim Col%: Col = Si(Fny)
Dim F
For Each F In Itr(Fny)
    If Not HasEle(InpColNy, F) Then
        WsC(Ws, Col).EntireColumn.Delete
    End If
    Col = Col - 1
Next
End Sub

Private Sub R6111_PasteAsVal(A As Worksheet)

End Sub
Private Sub PasteWsAsVal(Ws As Worksheet)

End Sub

Private Sub R5_GenFx()
Dim W As Database, OupFx$

End Sub
Private Sub Z_Rpt()
Dim WDb As Database, B As IGenr
GoSub ZZ
Exit Sub
ZZ:
    Set WDb = Nothing
    Set B = New ATaxExpCmp_OupTblGenr
    GoTo Tst
Tst:
    'Rpt_Oup WDb, B
    Return
End Sub
Private Sub R4_GenOupTbl()
A.Genr.GenOupTblFmTmpInp B.WDb
End Sub
