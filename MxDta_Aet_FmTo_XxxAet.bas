Attribute VB_Name = "MxDta_Aet_FmTo_XxxAet"
Option Compare Text
Const CMod$ = "MxDta_Aet_To."
Option Explicit

Function LnAet$(Aet As Dictionary):            LnAet = JnSpc(AvAet(Aet)): End Function
Function SyAet(Aet As Dictionary) As String(): SyAet = SyAy(Aet.Keys):    End Function
Function TmlAet$(Aet As Dictionary):      Stop '      TmlAet = TmlAy(AvAet(Aet)):   End Function
End Function
Function AvAet(Aet As Dictionary) As Variant(): AvAet = AvItr(Aet.Keys): End Function
Function ItmFstAet(Aet As Dictionary)
Dim I: For Each I In Aet.Keys
    Asg I, ItmFstAet: Exit Function
Next
End Function
