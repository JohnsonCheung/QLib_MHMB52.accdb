VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_FcTp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_FcTp."
Property Get Fn$():       Fn = "FcTp.xlsx":    End Property
Sub Edt():                EdtAttTpC Fn:        End Sub
Sub ImpIf():              ImpAttTpC Tp, Fn:    End Sub
Sub Exp(FfnTo$):          ExpAttTpC Fn, FfnTo: End Sub
Function Tp$():      Tp = Pth & Fn:            End Function
Function Pth$():    Pth = PthTp:               End Function
