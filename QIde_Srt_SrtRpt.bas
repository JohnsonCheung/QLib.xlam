Attribute VB_Name = "QIde_Srt_SrtRpt"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Srt_Rpt."
Private Const Asm$ = "QIde"
Private Function SrtRpt(Src$(), Optional Mdn$) As String()
Dim X As Dictionary
Dim Y As Dictionary
Set X = DiMthnqLines(Src, Mdn)
Set Y = SrtDic(X)
SrtRpt = FmtCmpgDic(X, Y, "BefSrt", "AftSrt")
End Function

Private Sub Z_SrtRpt()
Brw SrtRptzM(CMd)
End Sub

Property Get SrtRptM() As String()
SrtRptM = SrtRptzM(CMd)
End Property

Function SrtSrc(Src$()) As String()
SrtSrc = SplitCrLf(JnStrDic(SrtDic(DiMthnqLines(Src)), vb2CrLf))
End Function
Function SrtRptzP(P As VBProject) As String()
Dim O$(), C As VBComponent
For Each C In P.VBComponents
    PushIAy O, SrtRptzM(C.CodeModule)
Next
SrtRptzP = O
End Function

Function SrtRptDiczP(P As VBProject) As Dictionary
Dim C As VBComponent, O As New Dictionary, Md As CodeModule
    For Each C In P.VBComponents
        Set Md = C.CodeModule
        O.Add Mdn(Md), SrtRptzM(Md)
    Next
Set SrtRptDiczP = O
End Function

Function SrtRptzM(M As CodeModule) As String()
SrtRptzM = SrtRpt(Src(M), Mdn(M))
End Function

