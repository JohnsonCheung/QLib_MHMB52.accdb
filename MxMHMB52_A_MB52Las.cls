VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_MB52Las"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_MB52IOFun."
Private WIo As New MxMHMB52_A_MB52Io
Function Fxi$():                      Fxi = WIo.Fxi(MH.TbOH.YmdLas):                End Function
Function Fxo$():                      Fxo = WIo.Fxo(MH.TbOH.YmdLas):                End Function
Function FxiPthi$():              FxiPthi = EleMax(FxyiPthi):                       End Function
Function FxyiPthi() As String(): FxyiPthi = Ffny(WIo.Pthi, "MB52 20??-??-??.XLSX"): End Function
Sub OpnFxo():                               MaxvFx Fxo:                             End Sub
Sub OpnFxi():                               MaxvFx Fxi:                             End Sub
