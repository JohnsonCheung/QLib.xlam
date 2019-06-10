Attribute VB_Name = "QIde_Cmp_CntCmp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Cnt_Cmp."
Private Const Asm$ = "QIde"
Type CntgCmp
    NMod As Integer
    NCls As Integer
    NDoc As Integer
    NOth As Integer
    Locked As Boolean
End Type
Enum EmHdr
    EiNoHdr
    EiWiHdr
End Enum

Function CntgCmp(NMod%, NCls%, NDoc%, NOth%) As CntgCmp
With CntgCmp
    .NMod = NMod
    .NCls = NCls
    .NDoc = NDoc
    .NOth = NOth
End With
End Function

Sub CntCmpP()
CntCmpzP CPj
End Sub
Function Fny_CntgCmp() As String()

End Function

Function Dry_CntgCmp(P As VBProject) As Variant()

End Function

Function DCntgCmpzP(P As VBProject) As Drs
DCntgCmpzP = Drs(Fny_CntgCmp, Dry_CntgCmp(P))
End Function

Sub CntCmpzP(A As VBProject)
DmpRec DCntgCmpzP(A)
End Sub

Function NCmp%(A As CntgCmp)
With A
NCmp = .NMod + .NCls + .NDoc + .NOth
End With
End Function

Function CntgCmpLin$(A As CntgCmp, Optional Hdr As EmHdr = EiWiHdr)
Dim Pfx$
If Hdr = EiWiHdr Then Pfx = "Cmp Mod Cls Doc Oth" & vbCrLf
With A
CntgCmpLin = Pfx & NCmp(A) & " " & .NMod & " " & .NCls & " " & .NDoc & " " & .NOth
End With
End Function
Function CMP(Cmpn) As VBComponent
Set CMP = CPj.VBComponents(Cmpn)
End Function

Function NCls%()
NCls = NClszP(CPj)
End Function
Function NCmpPj%()
NCmpPj = CPj.VBComponents.Count
End Function


Function NModPj%()
NModPj = NModzP(CPj)
End Function
'===============================================

Function CntgCmpzP(P As VBProject) As CntgCmp
If P.Protection = vbext_pp_locked Then CntgCmpzP.Locked = True: Exit Function
CntgCmpzP = CntgCmp(NModzP(P), NClszP(P), NDoczP(P), NOthCmpzP(P))
End Function
'----------------------------------------------
Function CntgCmpP() As CntgCmp
CntgCmpP = CntgCmpzP(CPj)
End Function
'==============================================
Function NCmpzP%(P As VBProject)
If P.Protection = vbext_pp_locked Then Exit Function
NCmpzP = P.VBComponents.Count
End Function

Function NModzP%(P As VBProject)
NModzP = NCmpzTy(P, vbext_ct_StdModule)
End Function


Function NClszP%(P As VBProject)
NClszP = NCmpzTy(P, vbext_ct_ClassModule)
End Function


Function NCmpzTy%(P As VBProject, Ty As vbext_ComponentType)
If P.Protection = vbext_pp_locked Then Exit Function
Dim C As VBComponent
Dim O%
For Each C In P.VBComponents
    If C.Type = Ty Then O = O + 1
Next
NCmpzTy = O
End Function

Function NOthCmpzP%(P As VBProject)
NOthCmpzP = NCmpzP(P) - NClszP(P) - NModzP(P) - NDoczP(P)
End Function

