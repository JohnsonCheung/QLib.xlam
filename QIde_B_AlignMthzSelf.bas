Attribute VB_Name = "QIde_B_AlignMthzSelf"
Option Compare Text
Option Explicit

Sub AlignMthzSelf()
'Cpy Md
    Const TMdn$ = "QIde_B_AlignMth"      ' #The-Mdn
    Const TmMdn$ = "ATmp"                ' #Tmp-Mdn
:                                    EnsCls CPj, TmMdn
    Dim FmM As CodeModule: Set FmM = Md(TMdn)
    Dim ToM As CodeModule: Set ToM = Md(TmMdn)
    Dim OIsCpy As Boolean:  OIsCpy = CpyMd(FmM, ToM)
:                                    If OIsCpy Then MsgBox "Copied": Exit Sub

'Align
    Const TMthn$ = "AlignMthzLno"      ' #The-Mthn
    Dim M As CodeModule: Set M = Md(TMdn)
    Dim MthLno&:        MthLno = MthLnozMM(M, TMthn)
    'ATmp.AlignMthzLno M, MthLno, Upd:=EiUpdAndRpt, IsUpdSelf:=True
End Sub

Sub AlignMthErzLnk()
Dim M As CodeModule: Set M = Md("QDao_Lnk_ErzLnk")
Dim L&: L = MthLnozMM(M, "ErzLnk")
QIde_B_AlignMth.AlignMthzLno M, L
End Sub

