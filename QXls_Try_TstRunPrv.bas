Attribute VB_Name = "QXls_Try_TstRunPrv"
Option Compare Text
Option Explicit
Private Const CMod$ = "AA."
Sub XXX1A()
Run "XXX1B" 'XXX1B is private.  Running Private is Ok in Exl.  But Fun-Nm cannot be Exl Address.
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

'
