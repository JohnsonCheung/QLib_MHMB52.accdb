Attribute VB_Name = "MxIde_Dcl_Udt_Udty"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dcl_Udt_Udty."

Function Udty(Dcl$(), Udtn$) As String():             Udty = AwBei(Dcl, BeiUdt(Dcl, Udtn)): End Function
Function UdtyM(M As CodeModule, Udtn$) As String():  UdtyM = Udty(DclM(M), Udtn):           End Function
Function UdtyMC(Udtn$) As String():                 UdtyMC = UdtyM(CMd, Udtn):              End Function

Function UdtyyMC() As Variant():               UdtyyMC = UdtyyM(CMd):       End Function
Function UdtyyM(M As CodeModule) As Variant():  UdtyyM = UdtyyDcl(DclM(M)): End Function
Function UdtyyDcl(Dcl$()) As Variant()
Dim B() As Bei: B = BeiyUdt(Dcl)
Dim J&: For J = 0 To UbBei(B)
    PushI UdtyyDcl, AwBei(Dcl, B(J))
Next
End Function

Function UdtlUdtn$(Dcl$(), Udtn$): UdtlUdtn = JnCrLf(Udty(Dcl, Udtn)): End Function

Private Sub B_UdtyyPC():                                BrwLyy UdtyyPC:    End Sub
Function UdtyyPC() As Variant():              UdtyyPC = UdtyyP(CPj):       End Function
Function UdtyyP(P As VBProject) As Variant():  UdtyyP = UdtyyDcl(DclP(P)): End Function
