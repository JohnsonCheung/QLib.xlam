Attribute VB_Name = "BInpFils"
Option Explicit
Type InpFil: Fn As String: Ffn As String: End Type
Type InpFils: N As Integer: Ay() As InpFil: End Type
Private Function InpFil(Fn$, Ffn$) As InpFil
With InpFil: .Fn = Fn: .Ffn = Ffn: End With
End Function
Function InpFilzLin(InpFilLin$) As InpFil
With BrkSpc(InpFilLin)
InpFilzLin = InpFil(.S1, .S2)
End With
End Function
Function IsRmkLin(Lin$) As Boolean
Select Case FstChr(Lin)
Case "#", "@": IsRmkLin = True
End Select
End Function

Function RmvRmkLin(Sy$()) As String()
Dim L, Lin$
For Each L In Itr(Sy)
    Lin = L
    If Not IsRmkLin(Lin) Then
        PushS RmvRmkLin, Lin
    End If
Next
End Function
Sub PushInpFil(O As InpFils, M As InpFil)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub
Function InpFils(Src$()) As InpFils
Dim Lin$, I
For Each I In Itr(RmvRmkLin(Src))
    Lin = I
    PushInpFil InpFils, InpFilzLin(Lin)
Next
End Function


Function ErzInpFils(A As InpFils) As String()

End Function
Function SampSrczInpFil() As String()
Erase XX
X "MB52 C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipCost\Sample\MB52 2018-07-30.xls"
X "UOM  C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipCost\Sample\sales text.xlsx"
X "ZHT1 C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipCost\Sample\ZHT1.XLSX"
SampSrczInpFil = XX
Erase XX
End Function

