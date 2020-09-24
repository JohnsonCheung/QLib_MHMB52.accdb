Attribute VB_Name = "MxAcs_Acs_Acsa"
Option Compare Text
Option Explicit
Const CMod$ = "MxAcs_Acs_Acsa."
Private WAAcs As Access.Application
Function AAcs() As Access.Application
If Not IsAcsOk(WAAcs) Then Set WAAcs = New Access.Application: MinvAcs WAAcs
Set AAcs = WAAcs
End Function
Function AAcsFb(Fb, Optional IsExl As Boolean) As Access.Application
OpnAcsFb Fb, AAcs, IsExl
Set AAcsFb = WAAcs
End Function
