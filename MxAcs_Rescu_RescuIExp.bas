Attribute VB_Name = "MxAcs_Rescu_RescuIExp"
Option Compare Text
Option Explicit
Const CMod$ = "MxAcs_Acs_RescuFbByIExp."
Private WAcs As Access.Application, WIFb, WPthWrk$
Sub RescuByIExp(Pth$)
Set WAcs = New Access.Application
WPthWrk = ""
Dim WIFb: For Each WIFb In ItrFbCorrud(Pth)
    WExp
    WImp
Next
QuitAcs WAcs
End Sub
Private Sub WExp()
WAcs.OpenCurrentDatabase WIFb
ExpAcsC
'WAcs, WPthWrk
End Sub
Private Sub WImp()
Const CSub$ = CMod & "WExp"
Dim FbTo$: FbTo = WFbRescu
DltFfnIf FbTo
CrtFb FbTo
WAcs.OpenCurrentDatabase FbTo
ImpAcs WAcs, WPthWrk
WAcs.CloseCurrentDatabase
End Sub
Private Function WFbRescu$()
WFbRescu = FfnRplFnsfx(WIFb, "(corrupted)", "(rescured)")
End Function
