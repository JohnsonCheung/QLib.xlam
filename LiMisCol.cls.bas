VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiMisCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Ffn As String
Public T As String
Public Wsn As String
Public MisFset As New Aset
Public EptFset As New Aset
Public ActFset As New Aset
Friend Function Init(Ffn, T, EptFset As Aset, ActFset As Aset, Optional Wsn) As LiMisCol
With Me
    .Ffn = Ffn
    .T = T
    .Wsn = Wsn
    Set .EptFset = EptFset
    Set .ActFset = ActFset
    If EptFset.Minus(ActFset).Cnt = 0 Then Thw CSub, "No missing Fny, should not create this LiMisCol instance", "EptFset ActFset", EptFset.Sy, ActFset.Sy
End With
Set Init = Me
End Function

Property Get MisMsg() As String()
Dim N$: N = FmtQQ("Mis-Columns in-? in-? Actual-columns-in-table Expected-columns", FfnKd(Ffn), TblKd(Ffn))
Dim Mis$()
    Mis = EptFset.Minus(ActFset).Sy
Dim M$
    Select Case Sz(Mis)
    Case 1: M = "There is one column missing"
    Case Else: M = FmtQQ("There are ? columns missing", Sz(Mis))
    End Select
MisMsg = LyzMsgNap(M, N, Mis, Ffn, T, ActFset.Sy, EptFset.Sy)
End Property

