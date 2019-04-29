Attribute VB_Name = "MDta_Fmt"
Option Explicit
Sub VcDrs(A As Drs, Optional MaxColWdt% = 100, Optional BrkColnn$, Optional Fnn$)
BrwDrs A, MaxColWdt, BrkColnn, Fnn, UseVc:=True
End Sub

Sub BrwDrs(A As Drs, Optional MaxColWdt% = 100, Optional BrkColnn$, Optional Fnn$, Optional UseVc As Boolean)
BrwAy FmtDrs(A, MaxColWdt, BrkColnn), Fnn, UseVc
End Sub

Function FmtDrs(A As Drs, Optional MaxColWdt% = 100, Optional BrkColnn$, Optional ShwZer As Boolean, Optional HidIxCol As Boolean) As String()
'If BrkColNm changed, insert a break line if BrkColNm is given
Dim Drs As Drs
    Drs = DrsAddIxCol(A, HidIxCol)
Dim BrkColIxAy&()
    BrkColIxAy = IxAy(A.Fny, TermAy(BrkColnn))
Dim Dry()
    Dry = Drs.Dry
    PushI Dry, Drs.Fny

Dim Ay$()
    Ay = FmtDry(Dry, MaxColWdt, BrkColIxAy, ShwZer) '<== Will insert break line if BrkColIx>=0

Dim U&: U = UB(Ay)
Dim Hdr$: Hdr = Ay(U - 1)
Dim Lin$: Lin = Ay(U)
FmtDrs = AyeLasNEle(AyAdd(Sy(Lin, Hdr), Ay), 2)
PushI FmtDrs, Lin
End Function


Function FmtDt(A As Dt, Optional MaxColWdt% = 100, Optional BrkColNm$, Optional ShwZer As Boolean, Optional HidIxCol As Boolean) As String()
PushI FmtDt, "*Tbl " & A.DtNm
PushIAy FmtDt, FmtDrs(DrszDt(A), MaxColWdt, BrkColNm, ShwZer, HidIxCol)
End Function

Private Sub Z_FmtDrs()
Dim A As Drs, MaxColWdt%, BrkColVbl$, ShwZer As Boolean, HidIxCol As Boolean
A = SampDrs
GoSub Tst
Exit Sub
Tst:
    Act = FmtDrs(A, MaxColWdt, BrkColVbl, ShwZer, HidIxCol)
    Brw Act: Stop
    C
    Return
End Sub

Private Sub Z_FmtDt()
Dim A As Dt, MaxColWdt%, BrkColNm$, ShwZer As Boolean
'--
A = SampDt1
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
