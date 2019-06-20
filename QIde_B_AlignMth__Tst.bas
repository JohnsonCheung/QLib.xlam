Attribute VB_Name = "QIde_B_AlignMth__Tst"

Sub Z11()
Const Mdn$ = "QIde_B_AlignMth"
Const Mthn$ = "AlignMthDimzML"
Dim M As CodeModule: Set M = Md(Mdn)
Dim L&: L = MthLnozMM(M, Mthn)
QIde_B_AlignMth.AlignMthDimzML M, L, Rpt:=EiRptOnly
End Sub

Private Sub Z_AlignMthDimzML()
'Cpy Md
    Const TMdn$ = "QIde_B_AlignMth"      ' #The-Mdn
    Const TmMdn$ = "ATmp"                ' #Tmp-Mdn
    Dim OEns:                        EnsCls CPj, TmMdn
    Dim FmM As CodeModule: Set FmM = Md(TMdn)
    Dim ToM As CodeModule: Set ToM = Md(TmMdn)
    Dim OIsCpy As Boolean:  OIsCpy = CpyMd(FmM, ToM)
    Dim OExit:                       If OIsCpy Then MsgBox "Copied": Exit Sub

'Align
    Const TMthn$ = "AlignMthDimzML"      ' #The-Mthn
    Dim M As CodeModule: Set M = Md(TMdn)
    Dim MthLno&:        MthLno = MthLnozMM(M, TMthn)
    ATmp.AlignMthDimzML M, MthLno, Rpt:=EiUpdAndRpt, IsUpdSelf:=True
End Sub

Sub Z1()
Z_AlignMthDimzML
End Sub

Sub Z3()
Dim M As CodeModule: Set M = Md("QDao_Lnk_ErzLnk")
Dim L&: L = MthLnozMM(M, "ErzLnk")
QIde_B_AlignMth.AlignMthDimzML M, L
End Sub

