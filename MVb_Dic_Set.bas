Attribute VB_Name = "MVb_Dic_Set"
Option Explicit

Property Get EmpAset() As Aset
Set EmpAset = New Aset
End Property

Function CvAset(A) As Aset
Set CvAset = A
End Function

Function IsAset(A) As Boolean
IsAset = TypeName(A) = "Aset"
End Function

Function AsetzAp(ParamArray Ap()) As Aset
Dim Av(): Av = Ap
Set AsetzAp = AsetzAy(Av)
End Function
Function AsetzItr(Itr) As Aset
Set AsetzItr = EmpAset
AsetzItr.PushItr Itr
End Function
Function AsetzFF(FF) As Aset
Set AsetzFF = AsetzAy(NyzNN(FF))
End Function
Function AsetzAy(A) As Aset
Set AsetzAy = EmpAset
AsetzAy.PushAy A
End Function
Function AsetzSsl(Ssl) As Aset
Set AsetzSsl = AsetzAy(SySsl(Ssl))
End Function

