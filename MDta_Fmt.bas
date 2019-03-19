Attribute VB_Name = "MDta_Fmt"
Option Explicit
Sub VcDrs(A As Drs, Optional MaxColWdt% = 100, Optional BrkColNN, Optional Fnn$)
BrwDrs A, MaxColWdt, BrkColNN, Fnn, UseVc:=True
End Sub

Sub BrwDrs(A As Drs, Optional MaxColWdt% = 100, Optional BrkColNN, Optional Fnn$, Optional UseVc As Boolean)
BrwAy FmtDrs(A, MaxColWdt, BrkColNN), Fnn, UseVc
End Sub

Function FmtDrs(A As Drs, Optional MaxColWdt% = 100, Optional BrkColNN, Optional ShwZer As Boolean, Optional HidIxCol As Boolean) As String()
'If BrkColNm changed, insert a break line if BrkColNm is given
Dim Drs As Drs
    Set Drs = DrsAddIxCol(A, HidIxCol)
Dim BrkColIx%
    BrkColIx = IxzAy(A.Fny, BrkColNN)
Dim Dry()
    Dry = Drs.Dry
    PushI Dry, Drs.Fny

Dim Ay$()
    Ay = FmtDry(Dry, MaxColWdt, BrkColIx, ShwZer) '<== Will insert break line if BrkColIx>=0

Dim U&: U = UB(Ay)
Dim Hdr$: Hdr = Ay(U - 1)
Dim Lin$: Lin = Ay(U)
FmtDrs = AyeLasNEle(AyAdd(Sy(Lin, Hdr), Ay), 2)
PushI FmtDrs, Lin
End Function

Function FmtDs(A As Ds, Optional MaxColWdt% = 100, Optional ShwZer As Boolean, Optional HidIxCol As Boolean) As String()
PushI FmtDs, "*Ds " & A.DsNm
Dim I
For Each I In A.DtAy
    PushIAy FmtDs, FmtDt(CvDt(I), MaxColWdt, , ShwZer, HidIxCol)
Next
End Function

Function FmtDt(A As Dt, Optional MaxColWdt% = 100, Optional BrkColNm$, Optional ShwZer As Boolean, Optional HidIxCol As Boolean) As String()
PushI FmtDt, "*Tbl " & A.DtNm
PushIAy FmtDt, FmtDrs(DrszDt(A), MaxColWdt, BrkColNm, ShwZer, HidIxCol)
End Function

Private Sub Z_FmtDrs()
Dim A As Drs, MaxColWdt%, DtBrkLinMapStr$, NoIxCol As Boolean
Set A = SampDrs
GoSub Tst
Exit Sub
Tst:
    Act = FmtDrs(A, MaxColWdt, DtBrkLinMapStr, NoIxCol)
    'Brw Act: Stop
    C
    Return
End Sub

Private Sub Z_FmtDt()
Dim A As Dt, MaxColWdt%, BrkColNm$, ShwZer As Boolean
'--
Set A = SampDt1
'Ept = Z_DteTimStrpt1
GoSub Tst
'--
Exit Sub
Tst:
    Act = FmtDt(A, MaxColWdt, BrkColNm, ShwZer)
    C
    Return
End Sub

Private Sub ZZ()
Dim A As Drs
Dim B%
Dim C$
Dim D As Boolean
Dim E As Ds
Dim F As Dt
FmtDrs A, B, C, D, D
FmtDt F, B, C, D, D
End Sub

Private Sub Z()
Z_FmtDrs
'Z_FmtDt
End Sub
