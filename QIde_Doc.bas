Attribute VB_Name = "QIde_Doc"
Option Explicit
Private Const CMod$ = "MIde_Doc."
Private Const Asm$ = "QIde"
Function DocLy(DclLy$()) As String()
Dim Lin, N$
For Each Lin In Itr(DclLy)
    N = CnstNm(Lin)
    If Left(N, 5) = "Docz" Then PushI DocLy, Mid(N, 6) & " " & ValzCnstLin(Lin)
Next
End Function

Function NDocInPj%(A As VBProject)
NDocInPj = NCmpzTy(A, vbext_ct_Document)
End Function
Function IsDocNm(S$) As Boolean
If Not IsNm(S) Then Exit Function
IsDocNm = Left(S, 5) = "Docz"
End Function

Function DocDiczPj(A As VBProject) As Dictionary
Dim O As New Dictionary, Dcl
For Each Dcl In DclDiczPj(A).Items
    PushDic O, DocDiczDcl(CStr(Dcl))
Next
Set DocDiczPj = DicSrt(O)
End Function
Sub Doc(Nm$)
If DocDicInPj.Exists(Nm) Then D DocDicInPj(Nm) Else D "Not exist"
'#BNmMIS is Method-B-Nm-of-Missing.
'           Missing means the method is found in FmPj, but not ToPj
'#FmDicB is MthDic-of-MthBNm-zz-MthLines.   It comes from FmPj
'#ToDicA is MthDic-of-MthANm-zz-MthLinesAy. It comes from ToPj
'#ToDicAB is ToDicA and FmDicB
'#ANm is method-a-name, NNN or NNN:YYY
'        If the method is Sub|Fun, just MthNm
'        If the method is Prp    ,      MthNm:MthTy
'        It is from ToPj (#ToA)
'        One MthANm will have one or more MthLines
'#BNm is method-b-name, MMM.NNN or MMM.NNN:YYY
'        MdNm.MthNm[:MthTy]
'        It is from FmPj (#BFm)
'        One MthBNm will have only one MthLines
'#Missing is for each MthBNm found in FmDicB, but its MthNm is not found in any-method-name-in-ToDicA
'#Dif is for each MthBNm found in FmDicB and also its MthANm is found in ToDicA
'        and the MthB's MthLines is dif and any of the MthA's MthLines
'       (Note.MthANm will have one or more MthLines (due to in differmodule))
End Sub

