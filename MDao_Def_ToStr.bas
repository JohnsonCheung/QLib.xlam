Attribute VB_Name = "MDao_Def_ToStr"
Option Explicit
Const CMod$ = "MDao_Td_Str."

Function FdStrAyFds(A As Dao.Fields) As String()
Dim F As Dao.Field
For Each F In A
    PushI FdStrAyFds, FdStr(F)
Next
End Function

Function TdStr$(A As Dao.TableDef)
Dim T$, Id$, S$, R$
    T = A.Name
    If HasStdPkzTd(A) Then Id = "*Id"
    Dim Pk$(): Pk = Sy(T & "Id")
    Dim Sk$(): Sk = SkFnyzTd(A)
    If HasStdSkzTd(A) Then S = TLin(SyRpl(Sk, T, "*")) & " |"
    R = TLin(CvSy(AyMinusAp(FnyzTd(A), Pk, Sk)))
TdStr = JnSpc(SyzApNonBlank(T, Id, S, R))
End Function

Function FnyzTdLy(TdLy$()) As String()
Dim O$(), TdStr$, I
For Each I In Itr(TdLy)
    TdStr = I
    PushIAy O, FnyzTdLin(TdStr)
Next
FnyzTdLy = CvSy(AywDist(O))
End Function

Function TdStrzT$(A As Database, T$)
TdStrzT = TdStr(A.TableDefs(T))
End Function

Function SkFnyzTdLin(TdLin$) As String()
Dim A1$, T$, Rst$
    A1 = Bef(TdLin, "|")
    If A1 = "" Then Exit Function
AsgTRst A1, T, Rst
T = RmvSfx(T, "*")
Rst = Replace(Rst, "*", T)
SkFnyzTdLin = SySsl(Rst)
End Function

Function FdStr$(A As Dao.Field2)
Dim D$, R$, Z$, VTxt$, VRul, E$, S$
If A.Type = Dao.DataTypeEnum.dbText Then S = " TxtSz=" & A.Size
If A.DefaultValue <> "" Then D = "Dft=" & A.DefaultValue
If A.Required Then R = "Req"
If A.AllowZeroLength Then Z = "AlwZLen"
If A.Expression <> "" Then E = "Expr=" & A.Expression
If A.ValidationRule <> "" Then VRul = "VRul=" & A.ValidationRule
If A.ValidationText <> "" Then VTxt = "VTxt=" & A.ValidationText
FdStr = TLinzAp(A.Name, ShtTyzDao(A.Type), R, Z, VTxt, VRul, D, E, IIf((A.Attributes And Dao.FieldAttributeEnum.dbAutoIncrField) <> 0, "Auto", ""))
End Function


