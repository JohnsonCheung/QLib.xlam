Attribute VB_Name = "QDta_Samp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_Samp."
Private Const Asm$ = "QDta"
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
SampDrs1 = DrszFF("A B C", SampDry1)
End Property

Property Get SampDrs2() As Drs
SampDrs2 = DrszFF("A B C", SampDry2)
End Property

Property Get SampDrs() As Drs
SampDrs = DrszFF("A B C D E G H I J K", SampDry)
End Property

Property Get SampDFnyRs() As String()
SampDFnyRs = SyzSS("A B C D E F G")
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
AddDt SampDs, SampDt1
AddDt SampDs, SampDt2
SampDs.DsNm = "Ds"
End Property

Property Get SampDt1() As Dt
SampDt1 = DtzFF("SampDt1", "A B C", SampDry1)
End Property

Property Get SampDt2() As Dt
SampDt2 = DtzFF("SampDt2", "A B C", SampDry2)
End Property
