Attribute VB_Name = "QIde_Srt_SrtRpt"
Option Explicit
Private Const CMod$ = "MIde_Srt_Rpt."
Private Const Asm$ = "QIde"
Private Function SrtRpt(Src$()) As String()
Dim X As Dictionary
Dim Y As Dictionary
Set X = MthDic(Src)
Set Y = MthDic(SrcSrt(Src))
SrtRpt = FmtCmpDic(X, Y, "BefSrt", "AftSrt")
End Function

Private Sub Z_SrtRpt()
Brw SrtRpt(CurSrc)
End Sub

Property Get SrtRptMd() As String()
SrtRptMd = SrtRptzMd(CMd)
End Property

Function SrtRptzP(P As VBProject) As String()
Dim O$(), C As VBComponent
For Each C In P.VBComponents
    PushIAy O, SrtRptzMd(C.CodeModule)
Next
SrtRptzP = O
End Function

Function SrtRptDiczP(P As VBProject) As Dictionary
Dim C As VBComponent, O As New Dictionary, Md As CodeModule
    For Each C In P.VBComponents
        Set Md = C.CodeModule
        O.Add Mdn(Md), SrtRptzMd(Md)
    Next
Set SrtRptDiczP = O
End Function

Function SrtRptzMd(A As CodeModule) As String()
SrtRptzMd = SrtRpt(Src(A))
End Function

Function SrtDicMd(A As CodeModule) As Dictionary
Set SrtDicMd = AddDicKeyPfx(SSrcDic(Src(A)), Mdn(A) & ".")
End Function

