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

Property Get DoSamp1() As Drs
DoSamp1 = DrszFF("A B C", DyoSamp1)
End Property

Property Get DoSamp2() As Drs
DoSamp2 = DrszFF("A B C", DyoSamp2)
End Property

Property Get DoSamp() As Drs
DoSamp = DrszFF("A B C D E G H I J K", DyoSamp)
End Property

Property Get SampDFnyRs() As String()
SampDFnyRs = SyzSS("A B C D E F G")
End Property

Property Get DyoSamp1() As Variant()
DyoSamp1 = Array(SampDr1, SampDr2, SampDr3)
End Property

Property Get DyoSamp2() As Variant()
DyoSamp2 = Array(SampDr3, SampDr4, SampDr5)
End Property

Property Get DyoSamp() As Variant()
PushI DyoSamp, Array("A", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "A"))
PushI DyoSamp, Array("B", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "B"))
PushI DyoSamp, Array("C", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "C"))
PushI DyoSamp, Array("D", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "D"))
PushI DyoSamp, Array("E", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "E"))
PushI DyoSamp, Array("F", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "F"))
PushI DyoSamp, Array("G", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "G"))
End Property

Property Get SampDs() As Ds
AddDt SampDs, SampDt1
AddDt SampDs, SampDt2
SampDs.DsNm = "Ds"
End Property

Property Get SampDt1() As Dt
SampDt1 = DtzFF("SampDt1", "A B C", DyoSamp1)
End Property

Property Get SampDt2() As Dt
SampDt2 = DtzFF("SampDt2", "A B C", DyoSamp2)
End Property
