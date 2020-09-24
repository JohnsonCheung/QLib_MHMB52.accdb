VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_SkuPm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Text
Option Explicit
Private WFxi$, WPthCpyTo$, WIsCpyTo As Boolean
Private Function WWTool_Cdl$()
Const Macro$ = "Property Get {N}{Tyc}{As}(): {N} = W{N}: End Property"
Const Tmll$ = "[WMB52Pthi $] [WMB52PthCpy2 $] [WMB52PthCpy1 $] [WMB52IsCpy1 [] [ As Boolean]] [MB52IsCPy2 [] [ As Boolean]]" & _
" [WPhFxi $]" & _
" [WSkuFxi $] [WSkuPthCpyTo $] [WSkuIsCpyTo [] [ As Boolean]]" & _
" [WZHT0Pthi $] [ZHT0Fxi $] [WZHT0Fxw $]"
WWTool_Cdl = LinesMacroTmll(Macro, Tmll)
End Function
Property Get PthCpyTo$():          PthCpyTo = WPthCpyTo: End Property
Property Get Fxi$():               Fxi = WFxi:           End Property
Property Get IsCpyTo() As Boolean: IsCpyTo = WIsCpyTo:   End Property
Private Sub Class_Initialize()
RfhPm
End Sub

Sub RfhPm()
WFxi = PmvC("Sku_InpFx")
WIsCpyTo = PmvC("SkuLis_IsCpyTo")
WPthCpyTo = PmvC("SkuLis_CpyToPth")
End Sub
