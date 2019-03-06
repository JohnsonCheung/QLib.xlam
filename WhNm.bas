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
Private X_Re As RegExp
Dim X_ExlLikAy$()
Dim X_LikAy$()
Dim X_IsEmp As Boolean
Property Get IsEmp() As Boolean
IsEmp = X_IsEmp
End Property
Friend Function Init(Patn$, LikAy$(), ExlLikAy$()) As WhNm
If Patn = "" Then
    If Sz(LikAy) = 0 Then
        If Sz(ExlLikAy) = 0 Then
            X_IsEmp = True
            Set Init = Me
            Exit Function
        End If
    End If
    
End If
Set X_Re = Re
Set X_Re = RegExp(Patn)
X_ExlLikAy = ExlLikAy
X_LikAy = LikAy
Set Init = Me
End Function
Property Get Re() As RegExp
Set Re = X_Re
End Property
Property Get LikAy() As String()
LikAy = X_LikAy
End Property

Property Get ExlLikAy() As String()
ExlLikAy = X_ExlLikAy
End Property

Property Get ToStr$()
If IsEmp Then ToStr = "#Emp": Exit Function
Dim O$()
Push O, Quote(X_Re.Pattern, "Patn(*)")
Push O, Quote(TLin(X_LikAy), "LikAy(*)")
Push O, Quote(TLin(X_ExlLikAy), "ExlLikAy(*)")
ToStr = JnCrLf(O)
End Property
