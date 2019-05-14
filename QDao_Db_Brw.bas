Attribute VB_Name = "QDao_Db_Brw"
Option Explicit
Private Const CMod$ = "MDao_Db_Brw."
Private Const Asm$ = "QDao"

Sub BrwQ(A As Database, Q)
BrwDrs DrszQ(A, Q)
End Sub

