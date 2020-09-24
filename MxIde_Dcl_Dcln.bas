Attribute VB_Name = "MxIde_Dcl_Dcln"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dcl_Dcln."

Private Sub B_UdtnLn()
Debug.Assert UdtnLn("Type Udt") = "Udt"
Debug.Assert UdtnLn("Private Type Udt") = "Udt"
End Sub

Function EnmnLn$(L): EnmnLn = NmAftTm(RmvMdy(L), "Enum"): End Function
Function UdtnLn$(L): UdtnLn = NmAftTm(RmvMdy(L), "Type"): End Function
