VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_SHldTp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_SHldTp."

Sub Fmt()
Static X As New MxMHMB52_A_SHldTp_Fmt
X.Fmt
End Sub
Property Get Fn$():  Fn = "SHldTp.xlsx":  End Property
Sub Edt():           EdtAttTpC Fn:        End Sub
Sub ImpIf():         ImpAttTpC Tp, Fn:    End Sub
Sub Exp(FfnTo$):     ExpAttTpC Fn, FfnTo: End Sub
Property Get Tp$():  Tp = Pth & Fn:       End Property
Property Get Pth$(): Pth = PthTp:         End Property
