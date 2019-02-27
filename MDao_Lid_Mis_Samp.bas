Attribute VB_Name = "MDao_Lid_Mis_Samp"
Option Explicit
Private A$()
Function SampLidMis() As LidMis
Set SampLidMis = New LidMis
SampLidMis.Init Ffn, Tbl, Col, Ty
End Function
Private Function Ffn() As Aset
Set Ffn = New Aset
Ffn.PushAy Array("c:\sdlkf\lksdf\skldf", "c:\sdlkfj\sdklf\sdf", "c:\sdf\sdfk", "c:\sdf\sdfdf")
End Function
Private Function Tbl() As LidMisTbl()
PushObj Tbl, Tbl1
PushObj Tbl, Tbl2
PushObj Tbl, Tbl3
PushObj Tbl, Tbl4
End Function
Private Function Col() As LidMisCol()
PushObj Col, Col1
PushObj Col, Col2
PushObj Col, Col3
End Function
Private Function Ty() As LidMisTy()
PushObj Ty, Ty1
PushObj Ty, Ty2
PushObj Ty, Ty3
PushObj Ty, Ty4
End Function
Private Function Ty1() As LidMisTy
Set Ty1 = New LidMisTy
Ty1.Init "c:\sdfjldf", "sd", "Sheet1", Ty1Col
End Function
Private Function Ty2() As LidMisTy
Set Ty2 = New LidMisTy
Ty2.Init "c:\sdfjldf", "sd", "Sheet1", Ty2Col
End Function
Private Function Ty3() As LidMisTy
Set Ty3 = New LidMisTy
Ty3.Init "c:\sdfjldf", "sd", "Sheet1", Ty3Col
End Function
Private Function Ty4() As LidMisTy
Set Ty4 = New LidMisTy
Ty4.Init "c:\sdfjldf", "sd", "Sheet1", Ty4Col
End Function
Private Function Ty4Col() As LidMisTyc()
PushObj Ty4Col, Ty4Col1
PushObj Ty4Col, Ty4Col2
End Function
Private Function Ty3Col() As LidMisTyc()
PushObj Ty3Col, Ty3Col1
PushObj Ty3Col, Ty3Col2
End Function
Private Function Ty2Col() As LidMisTyc()
PushObj Ty2Col, Ty2Col1
PushObj Ty2Col, Ty2Col2
End Function
Private Function Ty1Col() As LidMisTyc()
PushObj Ty1Col, Ty1Col1
PushObj Ty1Col, Ty1Col2
End Function
Private Function Ty1Col1() As LidMisTyc
Set Ty1Col1 = New LidMisTyc
Ty1Col1.Init "lksdfj", "D", "DB"
End Function
Private Function Ty1Col2() As LidMisTyc
Set Ty1Col2 = New LidMisTyc
Ty1Col2.Init "lksdfj", "D", "DB"
End Function
Private Function Ty2Col1() As LidMisTyc
Set Ty2Col1 = New LidMisTyc
Ty2Col1.Init "lksdfj", "D", "DB"
End Function
Private Function Ty2Col2() As LidMisTyc
Set Ty2Col2 = New LidMisTyc
Ty2Col2.Init "lksdfj", "D", "DB"
End Function
Private Function Ty3Col1() As LidMisTyc
Set Ty3Col1 = New LidMisTyc
Ty3Col1.Init "lksdfj", "D", "DB"
End Function
Private Function Ty3Col2() As LidMisTyc
Set Ty3Col2 = New LidMisTyc
Ty3Col2.Init "lksdfj", "D", "DB"
End Function
Private Function Ty4Col1() As LidMisTyc
Set Ty4Col1 = New LidMisTyc
Ty4Col1.Init "lksdfj", "D", "DB"
End Function
Private Function Ty4Col2() As LidMisTyc
Set Ty4Col2 = New LidMisTyc
Ty4Col2.Init "lksdfj", "D", "DB"
End Function
Private Function Col1() As LidMisCol
Set Col1 = New LidMisCol
Col1.Init "C:\sdlkf\sdfk", "aaa", AsetzSsl("dsf sdf sdf"), AsetzSsl("skldf lsjkdf "), "Sheet1"
End Function
Private Function Col2() As LidMisCol
Set Col2 = New LidMisCol
Col2.Init "C:\sdlkf\sdfk", "aaa", AsetzSsl("dsf sdf sdf"), AsetzSsl("skldf lsjkdf "), "Sheet1"
End Function
Private Function Col3() As LidMisCol
Set Col3 = New LidMisCol
Col3.Init "C:\sdlkf\sdfk", "aaa", AsetzSsl("dsf sdf sdf"), AsetzSsl("skldf lsjkdf "), "Sheet1"
End Function

Private Function Tbl1() As LidMisTbl
Set Tbl1 = New LidMisTbl
Tbl1.Init "C:\kdsfjdf\sdf.xlsx", "MB52", "MB52", "Sheet1"
End Function
Private Function Tbl2() As LidMisTbl
Set Tbl2 = New LidMisTbl
Tbl1.Init "C:\kdsfjdf.accdb", "MB52", "MB52"
End Function
Private Function Tbl3() As LidMisTbl
Set Tbl3 = New LidMisTbl
Tbl1.Init "C:\kdsfjdf.xls", "MB52", "MB52", "Sheet1"
End Function
Private Function Tbl4() As LidMisTbl
Set Tbl4 = New LidMisTbl
Tbl1.Init "C:\kdsfjdf.mdb", "MB52", "MB52"
End Function

