Attribute VB_Name = "MDao_Db_Brw"
Option Explicit

Sub BrwQz(A As Database, Q)
BrwDrs DrszFbq(A, Q)
End Sub

Sub BrwQ(Q)
BrwQz CDb, Q
End Sub

