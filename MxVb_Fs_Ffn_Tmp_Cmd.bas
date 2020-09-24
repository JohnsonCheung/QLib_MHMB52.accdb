Attribute VB_Name = "MxVb_Fs_Ffn_Tmp_Cmd"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_Ffn_Tmp_Cmd."

Sub BrwPthCmd(): BrwPth WPthCmd: End Sub
Function FcmdCrt$(PfxCmd$, CdlCmd$)
FcmdCrt = WFfn(PfxCmd)
WrtStr CdlCmd, FcmdCrt
End Function

Private Function WFfn$(PfxCmd$): WFfn = WPthCmd & PfxCmd & "_" & StrNow15 & ".cmd": End Function
Private Function WPthCmd$()
Static X$: If X = "" Then X = PthAddFdrEns(PthTmp, ".Cmd")
WPthCmd = X
End Function
Sub ClrPthCmd(): ClrPth WPthCmd: End Sub
