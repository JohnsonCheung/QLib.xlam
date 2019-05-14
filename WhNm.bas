VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WhNm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const CMod$ = "WhNm."
Private X_Re As RegExp
Dim X_ExlLikSy$()
Dim X_LikeAy$()
Dim X_IsEmp As Boolean
Property Get IsEmp() As Boolean
IsEmp = X_IsEmp
End Property
Friend Function Init(Patn$, LikeAy$(), ExlLikSy$()) As WhNm
If Patn = "" Then
    If Si(LikeAy) = 0 Then
        If Si(ExlLikSy) = 0 Then
            X_IsEmp = True
            Set Init = Me
            Exit Function
        End If
    End If
    
End If
Set X_Re = Re
Set X_Re = RegExp(Patn)
X_ExlLikSy = ExlLikSy
X_LikeAy = LikeAy
Set Init = Me
End Function
Property Get Re() As RegExp
Set Re = X_Re
End Property

Property Get LikeAy() As String()
LikeAy = X_LikeAy
End Property

Property Get ExlLikSy() As String()
ExlLikSy = X_ExlLikSy
End Property

Property Get ToStr$()
If IsEmp Then ToStr = "#Emp": Exit Function
Dim O$()
Push O, Quote(X_Re.Pattern, "Patn(*)")
Push O, Quote(TLin(X_LikeAy), "LikeAy(*)")
Push O, Quote(TLin(X_ExlLikSy), "ExlLikSy(*)")
ToStr = JnCrLf(O)
End Property
