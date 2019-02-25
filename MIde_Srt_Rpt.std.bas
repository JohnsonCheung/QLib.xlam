Attribute VB_Name = "MIde_Srt_Rpt"
Option Explicit
Private Function SrtRptz(Src$()) As String()
Dim X As Dictionary
Dim Y As Dictionary
Set X = MthDic(Src)
Set Y = MthDic(SrtedSrcz(Src))
SrtRptz = FmtCmpDic(X, Y, "BefSrt", "AftSrt")
End Function

Private Sub Z_SrtRptz()
Brw SrtRptz(SrcMd)
End Sub

Property Get SrtRpt() As String()
SrtRpt = SrtRptzMd(CurMd)
End Property

Function SrtRptzPj(A As VBProject) As String()
Dim O$(), C As VBComponent
For Each C In A.VBComponents
    PushIAy O, SrtRptzMd(C.CodeModule)
Next
SrtRptzPj = O
End Function

Function SrtRptDiczPj(A As VBProject) As Dictionary
Dim C As VBComponent, O As New Dictionary, Md As CodeModule
    For Each C In A.VBComponents
        Set Md = C.CodeModule
        O.Add MdNm(Md), SrtRptzMd(Md)
    Next
Set SrtRptDiczPj = O
End Function

Function SrtRptzMd(A As CodeModule) As String()
SrtRptzMd = SrtRptz(Src(A))
End Function

Function SrtDicMd(A As CodeModule) As Dictionary
Set SrtDicMd = AddDicKeyPfx(SrtedSrcDic(Src(A)), MdNm(A) & ".")
End Function

