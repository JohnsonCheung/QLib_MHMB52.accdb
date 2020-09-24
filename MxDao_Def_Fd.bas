Attribute VB_Name = "MxDao_Def_Fd"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Def_Fd."

Function CvFd(V) As Dao.Field:    Set CvFd = V: End Function
Function CvFd2(V) As Dao.Field2: Set CvFd2 = V: End Function

Function CloneFd(A As Dao.Field2, Fldn) As Dao.Field2
Set CloneFd = New Dao.Field
With CloneFd
    .Name = Fldn
    .Type = A.Type
    .AllowZeroLength = A.AllowZeroLength
    .Attributes = A.Attributes
    .DefaultValue = A.DefaultValue
    .Expression = A.Expression
    .Required = A.Required
    .ValidationRule = A.ValidationRule
    .ValidationText = A.ValidationText
End With
End Function

Function IsEqFd(A As Dao.Field2, B As Dao.Field2) As Boolean
With A
    If .Name <> B.Name Then Exit Function
    If .Type <> B.Type Then Exit Function
    If .Required <> B.Required Then Exit Function
    If .AllowZeroLength <> B.AllowZeroLength Then Exit Function
    If .DefaultValue <> B.DefaultValue Then Exit Function
    If .ValidationRule <> B.ValidationRule Then Exit Function
    If .ValidationText <> B.ValidationText Then Exit Function
    If .Expression <> B.Expression Then Exit Function
    If .Attributes <> B.Attributes Then Exit Function
    If .Size <> B.Size Then Exit Function
End With
IsEqFd = True
End Function

Function Fdv(A As Dao.Field)
On Error Resume Next
Fdv = A.Value
End Function
