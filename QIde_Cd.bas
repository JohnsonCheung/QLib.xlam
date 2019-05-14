Attribute VB_Name = "QIde_Cd"
Public Const DoczPushgCd$ = "It a cd of a given tyn-X with these fun{PushX AddX PushXs} and this ty-Xs"
Public Const DoczCd$ = "Type:MthnSfx.  If a fun is XXXCd, it means its generating some vb fun?ty."
Private Type A
    Tyn As String
End Type
Private A As A
Function PushgCd(Tyn) As String
A.Tyn = Tyn
PushgCd = CdTys & CdPush & CdPushs & CdAdd
End Function
Private Function CdTys() As String

End Function
Private Function CdPushs() As String

End Function
Private Function CdPush() As String

End Function
Private Function CdAdd() As String

End Function

