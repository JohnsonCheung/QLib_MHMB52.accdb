Attribute VB_Name = "MxIde_Md_Mdn_MdfNm"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Md_Mdn_MdfNm."
Type MdfNm
    IsPrv As Boolean
    Nm As String
End Type
Function MdfNm(IsPrv As Boolean, Nm$) As MdfNm
With MdfNm
    .IsPrv = IsPrv
    .Nm = Nm
End With
End Function
