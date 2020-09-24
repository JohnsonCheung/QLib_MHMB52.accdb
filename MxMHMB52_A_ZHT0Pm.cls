VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_ZHT0Pm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Text
Option Explicit
Private WPthi$
Property Get PmnFxi$():             PmnFxi = "ZHT0_InpFx":                                    End Property
Property Get Pthi$():               Pthi = WPthi:                                             End Property
Property Get Fxi$():                Fxi = MH.ZHT0Pm.Fxi:                                      End Property
Function FxiSelSet$():  FxiSelSet = PmFxSelSetC(PmnFxi, "Select the [Sales Text.xlsx] file"): End Function
Private Sub Class_Initialize()
WPthi = MH.MB52IO.Pthi
End Sub
