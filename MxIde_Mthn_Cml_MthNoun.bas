Attribute VB_Name = "MxIde_Mthn_Cml_MthNoun"
Option Compare Text
Option Explicit
Sub VcMthnoun(): VcAy AySrtQ(AwDis(MthnounyPubPC)): End Sub
Function MthnounyPubPC() As String():
Dim A$(): A = MthiyNavvMthlnPubPC
Dim B$(): B = Tm1y(A)
Dim C$(): C = AeEle(B, ".")
MthnounyPubPC = AwDis(B)
End Function
