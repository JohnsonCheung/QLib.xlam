Attribute VB_Name = "MDta_Samp"
Option Explicit
Property Get SampDr1() As Variant()
SampDr1 = Array(1, 2, 3)
End Property

Property Get SampDr2() As Variant()
SampDr2 = Array(2, 3, 4)
End Property

Property Get SampDr3() As Variant()
SampDr3 = Array(3, 4, 5)
End Property

Property Get SampDr4() As Variant()
SampDr4 = Array(43, 44, 45)
End Property

Property Get SampDr5() As Variant()
SampDr5 = Array(53, 54, 55)
End Property

Property Get SampDr6() As Variant()
SampDr6 = Array(63, 64, 65)
End Property

Property Get SampDrs1() As Drs
Set SampDrs1 = Drs("A B C", SampDry1)
End Property

Property Get SampDrs2() As Drs
Set SampDrs2 = Drs("A B C", SampDry2)
End Property

Property Get SampDrs() As Drs
Set SampDrs = Drs("A B C D E G H I J K", SampDry)
End Property

Property Get SampDFnyRs() As String()
SampDFnyRs = SySsl("A B C D E F G")
End Property

Property Get SampDry1() As Variant()
SampDry1 = Array(SampDr1, SampDr2, SampDr3)
End Property

Property Get SampDry2() As Variant()
SampDry2 = Array(SampDr3, SampDr4, SampDr5)
End Property

Property Get SampDry() As Variant()
PushI SampDry, Array("A", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "A"))
PushI SampDry, Array("B", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "B"))
PushI SampDry, Array("C", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "C"))
PushI SampDry, Array("D", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "D"))
PushI SampDry, Array("E", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "E"))
PushI SampDry, Array("F", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "F"))
PushI SampDry, Array("G", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "G"))
End Property

Property Get SampDs() As Ds
'Set SampDs = Ds(DtAy(SampDt1, SampDt2), "SampDs")
End Property

Property Get SampDt1() As Dt
Set SampDt1 = Dt("SampDt1", "A B C", SampDry1)
End Property

Property Get SampDt2() As Dt
Set SampDt2 = Dt("SampDt2", "A B C", SampDry2)
End Property