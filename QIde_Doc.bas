Attribute VB_Name = "QIde_Doc"
Option Explicit
Private Const CMod$ = "MIde_Doc."
Private Const Asm$ = "QIde"

Function DocLy(DclLy$()) As String()
Dim Lin, N$, I
For Each I In Itr(DclLy)
    Lin = I
    N = Cnstn(Lin)
    If Left(N, 5) = "Docz" Then PushI DocLy, Mid(N, 6) & " " & StrValzCnstLin(Lin)
Next
End Function
Function NDoczP%(P As VBProject)
NDoczP = NCmpzTy(P, vbext_ct_Document)
End Function
Function IsDocNm(S) As Boolean
If Not IsNm(S) Then Exit Function
IsDocNm = Left(S, 5) = "Docz"
End Function

Function DocDiczP(P As VBProject) As Dictionary
Dim O As New Dictionary, Dcl
For Each Dcl In DclDiczP(P).Items
    PushDic O, DocDiczDcl(CStr(Dcl))
Next
Set DocDiczP = SrtDic(O)
End Function
Sub Doc(Nm)
If DocDicP.Exists(Nm) Then D DocDicP(Nm) Else D "Not exist"
'#BNmMIS is Method-B-Nm-of-Missing.
'           Missing means the method is found in FmPj, but not ToPj
'#FmDicB is MthDic-of-MthBNm-zz-MthLines.   It comes from FmPj
'#ToDicA is MthDic-of-MthANm-zz-MthLinesAy. It comes from ToPj
'#ToDicAB is ToDicA and FmDicB
'#ANm is method-a-name, NNN or NNN:YYY
'        If the method is Sub|Fun, just Mthn
'        If the method is Prp    ,      Mthn:MthTy
'        It is from ToPj (#ToA)
'        One MthANm will have one or more MthLines
'#BNm is method-b-name, MMM.NNN or MMM.NNN:YYY
'        Mdn.Mthn[:MthTy]
'        It is from FmPj (#BFm)
'        One MthBNm will have only one MthLines
'#Missing is for each MthBNm found in FmDicB, but its Mthn is not found in any-method-name-in-ToDicA
'#Dif is for each MthBNm found in FmDicB and also its MthANm is found in ToDicA
'        and the MthB's MthLines is dif and any of the MthA's MthLines
'       (Note.MthANm will have one or more MthLines (due to in differmodule))
End Sub

