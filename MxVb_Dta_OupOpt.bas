Attribute VB_Name = "MxVb_Dta_OupOpt"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_OupOpt."
Enum eOup: eOupDmp: eOupBrw: eOupVc: End Enum
Public Const EnmmOup$ = "eOup? Dmp Brw Vc"
Type TOup: PfxFn As String: Oup As eOup: End Type
Function TOup(PfxFn$, Oup As eOup) As TOup
With TOup
    .PfxFn = PfxFn
    .Oup = Oup
End With
End Function

Function TOupDmp() As TOup:                        TOupDmp = TOup("", eOupDmp):    End Function
Function TOupVc(Optional PfxFn$ = "Vc") As TOup:    TOupVc = TOup(PfxFn, eOupVc):  End Function
Function TOupBrw(Optional PfxFn$ = "Brw") As TOup: TOupBrw = TOup(PfxFn, eOupBrw): End Function
