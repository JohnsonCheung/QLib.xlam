VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MthNm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Nm$, MthTy$, MthMdy$
Friend Function Init(MthMdy, MthTy, Nm) As MthNm3
With Me
    .Nm = Nm
    .MthMdy = MthMdy
    .MthTy = MthTy
End With
Set Init = Me
End Function
Function Lin$(Optional Hdr As eHdr)
Dim Pfx$: If Hdr = eeWithHdr Then Pfx = "Mdy Ty Nm: "
Lin = Pfx & Apd(MthMdy, " ") & MthTy & " " & Nm
End Function
Property Get DNm$()
If IsEmp Then Exit Property
DNm = JnDotAp(Nm, ShtTy, ShtMdy)
End Property
Property Get IsEmp() As Boolean:  IsEmp = Nm = "":                              End Property
Property Get ShtMdy$():          ShtMdy = ShtMthMdy(MthMdy):                    End Property
Property Get ShtTy$():            ShtTy = ShtMthTy(MthTy):                      End Property
Property Get ShtKd$():            ShtKd = ShtMthKd(T1(MthTy)):                  End Property
Property Get MthKd$():            MthKd = T1(MthTy):                            End Property
Property Get IsPub() As Boolean:  IsPub = (MthMdy = "") Or (MthMdy = "Public"): End Property
Property Get IsPrv() As Boolean:  IsPrv = MthMdy = "Private":                   End Property
Property Get IsFrd() As Boolean:  IsFrd = MthMdy = "Friend":                    End Property
Property Get IsSub() As Boolean:  IsSub = MthTy = "Sub":                        End Property
Property Get IsPrp() As Boolean:  IsPrp = MthKd = "Property":                   End Property
Property Get IsFun() As Boolean:  IsFun = MthTy = "Function":                   End Property

Property Get MthNmDr() As String()
If Nm <> "" Then MthNmDr = Sy(Nm, ShtTy, ShtMdy)
End Property

