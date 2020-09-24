VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_GitIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_GitIO."

Property Get WsnFxi$():                WsnFxi = "Sheet1":                                            End Property
Property Get Pthi$():                  Pthi = MH.MB52Pm.Pthi:                                        End Property
Function Fxi$(D As Ymd):         Fxi = Pthi & WFn(D):                                                End Function
Private Function WFn$(D As Ymd): WFn = "Git " & Hypym(D.Y, D.M) & ".xlsx":                           End Function
Property Get YpStk%():                 YpStk = ValQC("Select YpStk from YpStk where NmYpStk='GIT'"): End Property
