Attribute VB_Name = "MxVb_Dta_Opt"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Opt."
Type Opt
    Som As Boolean
    Itm As Variant
End Type
Sub PushIOpt(OAy, M As Opt)
If M.Som Then PushI OAy, M.Itm
End Sub

Function SomItm(Itm) As Opt
SomItm.Som = True
SomItm.Itm = Itm
End Function
