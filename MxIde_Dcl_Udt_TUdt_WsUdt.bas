Attribute VB_Name = "MxIde_Dcl_Udt_TUdt_WsUdt"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dcl_Udt_UdtWs."

Function WsTUdtPC() As Worksheet:                              Set WsTUdtPC = WsTUdtP(CPj):                  End Function
Function WsTUdtP(P As VBProject) As Worksheet:                  Set WsTUdtP = WsTUdtFmt(WsDrs(DrstUdtP(P))): End Function
Private Function WsTUdtFmt(WsTUdt As Worksheet) As Worksheet: Set WsTUdtFmt = WsTUdt:                        End Function
