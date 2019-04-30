Attribute VB_Name = "MVb_Tst_TstHomClr"
Option Explicit

Sub ClsTstHom() ' Rmv-Empty-Pth Rmk-Pth-As-At
RmvEmpPthR TstHom
Ren_PthPj_AsAt
Ren_MdPth_AsAt
Ren_MthPth_AsAt
Ren_CasPth_AsAt
End Sub
Private Sub Ren_PthPj_AsAt()

End Sub
Private Sub Ren_MdPth_AsAt()

End Sub
Private Sub Ren_MthPth_AsAt()

End Sub
Private Sub Ren_CasPth_AsAt()
Ren CasPthSy
End Sub
Private Property Get CasPthSy() As String()

End Property
Private Sub Ren(PthSy)
Dim I
For Each I In Itr(PthSy)
    'RenPthAddPfx I, "@"
Next
End Sub
