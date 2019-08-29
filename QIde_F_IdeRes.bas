Attribute VB_Name = "QIde_F_IdeRes"
Option Explicit
Option Compare Text

Function ResLyzMthn(M As CodeModule, Mthn$) As String()
Dim Z$
    Z = MthLzM(M, Mthn)
    Stop
    If Z = "" Then
        Thw CSub, "Mthn not found", "Mthn Md", Mthn, Mdn(M)
        Exit Function
    End If
    Z = AeFstEle(Z)
    Z = AeLasEle(Z)
ResLyzMthn = RmvFstChrzAy(Z)
End Function

Function ReszMthn$(M As CodeModule, Mthn$)
':MthQn: :Dn|Nm ! if :Dn, Mdn.Mthn, If :Nm Mthn
':Dn:    :Nm.Nm #Dot-Nm#
':DDn:   :Nm{.Nm} #Dot-Dot-Nm#
ReszMthn = JnCrLf(ResLyzMthn(M, Mthn))
End Function


'
