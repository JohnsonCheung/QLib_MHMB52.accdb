Attribute VB_Name = "MxIde_Dcl_Udt_TUdt_zIntl_Udtny"
Option Compare Database
Option Explicit

Function UdtnyTUdty(U() As TUdt) As String()
Dim J%: For J = 0 To UbTUdt(U)
    PushI UdtnyTUdty, U(J).Udtn
Next
End Function
Function UdtmnyTUdty(U() As TUdt) As String()
Dim J%: For J = 0 To UbTUdt(U)
    PushI UdtmnyTUdty, Udtmn(U(J))
Next
End Function

Function FmtUdtilnyTUdty(U() As TUdt) As String(): FmtUdtilnyTUdty = FmtT3ry(UdtilnyTUdty(U)): End Function
Function UdtilnyTUdty(U() As TUdt) As String()
Dim J%: For J = 0 To UbTUdt(U)
    PushI UdtilnyTUdty, Udtiln(U(J))
Next
End Function
Private Function Udtmn$(U As TUdt)
With U
    Udtmn = IIf(.IsPrv, "Prv", "Pub") & " " & .Mdn & " " & .Udtn
End With
End Function
Private Function Udtiln$(U As TUdt): Udtiln = Udtmn(U) & " " & UmbnnTUdt(U): End Function

