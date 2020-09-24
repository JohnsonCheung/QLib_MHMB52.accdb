Attribute VB_Name = "MxIde_Mthn_Mthny_WithDash"
Option Compare Text
Const CMod$ = "MxIde_Mthn_Mthny_WithDash."
Option Explicit
Private Sub B_FunnyWithDashPC(): VcAy FunnyWithDashPC: End Sub
Function FunnyWithDashPC() As String()
Dim O$(): O = AePfx(AwSsub(FunnyPubPC, "_"), "B_")
O = AeSfx(O, "_Click")
O = AeSfx(O, "__Tst")
FunnyWithDashPC = SySrtQ(O)
End Function
