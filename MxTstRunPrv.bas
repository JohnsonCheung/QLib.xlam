Attribute VB_Name = "MxTstRunPrv"
Option Compare Text
Option Explicit
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxTstRunPrv."
Sub XXX1A()
Run "XXX1B" 'XXX1B is private.  Running Private is Ok in Xls.  But Fun-Nm cannot be Xls Address.
'In Acs, Running Private is not OK!!
End Sub
Private Sub XXX1B()
MsgBox "XXX1B"
End Sub
Private Function VerbPatn$(XX)
MsgBox XX
End Function
Sub F() 'Run "F" will fail
End Sub
