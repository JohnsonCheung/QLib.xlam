Attribute VB_Name = "QVb_Fs_DInp"
Option Compare Text
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_Fs_Ffn_MisEr."
Function ErzColDup(A As Drs, C$) As String()
Dim B As Drs: B = ColDup(A, C)
Dim Msg$: Msg = "Dup [" & C & "]"
ErzColDup = ErzDrsMsg(B, Msg)
End Function

Function DInpzSamp() As Drs
Erase XX
X "MB52 C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipCost\Sample\MB52 2018-07-30.xls"
X "UOM  C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipCost\Sample\sales text.xlsx"
X "ZHT1 C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipCost\Sample\ZHT1.XLSX"
DInpzSamp = DrszTRst("Inpn Ffn", XX)
End Function
Function ErzDrsMsg(A As Drs, Msg$) As String()
If NoReczDrs(A) Then Exit Function
Erase XX
XLin Msg
XDrs A
XLin
ErzDrsMsg = XX
End Function
Sub Z_ErzDInp()
Brw ErzDInp(DInpzSamp)
End Sub
Function ErzDInp(DInp As Drs) As String()
Dim E1$(), E2$(), E3$()
E1 = ErzColDup(DInp, "Ffn")
E2 = ErzColDup(DInp, "Inpn")
E3 = ErzFfnMis(DInp)
ErzDInp = Sy(E1, E2, E3)
End Function
Private Function ErzFfnMis(A As Drs) As String()
Dim I%: I = IxzAy(A.Fny, "Ffn")
Dim Dr, Dry(): For Each Dr In Itr(A.Dry)
    If Not HasFfn(Dr(I)) Then PushI Dry, Dr
Next
Dim B As Drs: B = DrszNewDry(A, Dry)
ErzFfnMis = ErzDrsMsg(B, "File not exist")
End Function
