VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const CMod$ = CLib & "UserForm1."
Private Type A
    Ix As Integer
    MdNy() As String
End Type
Private A As A
Property Get MdNy() As String()
If Si(A.MdNy) = 0 Then
    A.MdNy = Itn(CPj.VBComponents)
    A.Ix = 1
End If
MdNy = A.MdNy
End Property
Property Get NxtMdn$()
With A
    NxtMdn = MdNy()(.Ix)
    .Ix = .Ix + 1
End With
End Property
Sub CmdNxt_Click()
JmpMdn NxtMdn
End Sub

