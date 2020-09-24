Attribute VB_Name = "MxIde_Dcl_Udt_TUdt_TUdtPrp"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dcl_Udt_TUdt_TUdtPrp."

Function MdnyTUdty(U() As TUdt) As String()
Dim J%: For J = 0 To UbTUdt(U)
    PushI MdnyTUdty, U(J).Mdn
Next
End Function
Function UdtnyTUdty(U() As TUdt) As String()
Dim J%: For J = 0 To UbTUdt(U)
    PushI UdtnyTUdty, U(J).Udtn
Next
End Function

Function UmbnyTUdt(U As TUdt) As String()
Dim M() As TUmb: M = U.Mbr
Dim J%: For J = 0 To UbTUmb(M)
    PushI UmbnyTUdt, M(J).Mbn
Next
End Function
Function UmbnnTUdt$(U As TUdt): UmbnnTUdt = JnSpc(UmbnyTUdt(U)): End Function
Function MbnyTUdt(U As TUdt) As String()
Dim M() As TUmb: M = U.Mbr
Dim J%: For J = 0 To UbTUmb(M)
    PushI MbnyTUdt, M(J).Mbn
Next
End Function

