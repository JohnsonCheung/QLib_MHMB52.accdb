VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_PhPm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Text
Option Explicit
Private WFxi$
Property Get Fxi$():            Fxi = WFxi:              End Property
Private Sub Class_Initialize(): WFxi = PmvC("PH_InpFx"): End Sub
