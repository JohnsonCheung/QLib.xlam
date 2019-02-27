Attribute VB_Name = "MDta_Fmt"
Option Explicit
Sub BrwDrs(A As DRs, Optional MaxColWdt% = 100, Optional BrkColNm$, Optional Fnn$)
BrwAy FmtDrs(A, MaxColWdt, BrkColNm$), Fnn
End Sub

Function FmtDrs(A As DRs, Optional MaxColWdt% = 100, Optional BrkColNm$, Optional ShwZer As Boolean, Optional HidIxCol As Boolean) As String()
'If BrkColNm changed, insert a break line if BrkColNm is given
Dim DRs As DRs
    Set DRs = DrsAddIxCol(A, HidIxCol)
Dim BrkColIx%
    BrkColIx = IxzAy(A.Fny, BrkColNm)
Dim Dry()
    Dry = DRs.Dry
    PushI Dry, DRs.Fny

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

Function FmtDt(A As DT, Optional MaxColWdt% = 100, Optional BrkColNm$, Optional ShwZer As Boolean, Optional HidIxCol As Boolean) As String()
PushI FmtDt, "*Tbl " & A.DtNm
PushIAy FmtDt, FmtDrs(DrszDt(A), MaxColWdt, BrkColNm, ShwZer, HidIxCol)
End Function

Private Sub Z_FmtDrs()
Dim A As DRs, MaxColWdt%, DtBrkLinMapStr$, NoIxCol As Boolean
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
Dim A As DT, MaxColWdt%, BrkColNm$, ShwZer As Boolean
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
Dim A As DRs
Dim B%
Dim C$
Dim D As Boolean
Dim E As Ds
Dim F As DT
FmtDrs A, B, C, D, D
FmtDt F, B, C, D, D
End Sub

Private Sub Z()
Z_FmtDrs
'Z_FmtDt
End Sub
