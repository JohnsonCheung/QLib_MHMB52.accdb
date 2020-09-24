VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_MB52Io"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_MB52Io."
Private WPm As MxMHMB52_A_MB52Pm
Sub DoBrwPtho():                           BrwPth Ptho:                                   End Sub
Function Ptho$():                   Ptho = PthOup:                                        End Function
Function Fxi$(D As Ymd):             Fxi = Pthi & WFn(D):                                 End Function
Function Fxo$(D As Ymd):             Fxo = Ptho & W2Fn(D):                                End Function
Function Fxo1$(D As Ymd):           Fxo1 = WPm.PthCpy1 & W2Fn(D):                         End Function
Function Fxo2$(D As Ymd):           Fxo2 = WPm.PthCpy2 & W2Fn(D):                         End Function
Function FxyOup() As String():    FxyOup = Fxy(Ptho):                                     End Function
Private Function W2Fn$(D As Ymd):   W2Fn = "On Hand (MB52) " & Hypym(D.Y, D.M) & ".xlsx": End Function
Private Function WFn$(D As Ymd):     WFn = "MB52 " & Hypym(D.Y, D.M) & ".xlsx":           End Function
Property Get WsnFxi$():                    WsnFxi = "Sheet1":                             End Property
Property Get Pthi$():                      Pthi = WPm.Pthi:                               End Property


Function FmllnyWsDta() As String()
Dim O$()
'Unit
PushS O, "Btl/Unit =[@[Btl/AC]]/[@[Unit/AC]]"

'Siz
PushS O, "ml/Btl   =[@[Litre/Btl]]*1000"
PushS O, "Btl/AC'  =[@[Btl/AC]]"
PushS O, "Litre/SC =[@[Litre/Btl]] * [@[Btl/AC]] / [@[Unit/AC]] * [@[Unit/SC]]"

'OH
PushS O, "AC      =[@Btl] / [@[Btl/AC]]"
PushS O, "SC      =[@AC] * [@[Unit/AC]] / [@[Unit/SC]]"

'Pri
PushS O, "BtlUPr  =[@Val]/[Btl]"
PushS O, "AcUPr   =[@Val]/[@AC]"
PushS O, "ScUPr   =[@Val]/[@SC]"

'TaxItm
'Taxed
PushS O, "TaxItm = IF(OR(NOT(ISBLANK([@BchRat])),NOT(ISBLANK([@RatZHT0]))),""Y"","""")"
PushS O, "Taxed = IF(AND([@NoTax]<>""Y"",OR([@TaxItm]=""Y"",[@[3p]]=""Y""),[@TaxLoc]=""Y""),""Y"","""")"

'Amt
PushS O, "BchAmt  =IF([@Taxed]=""Y"",[@BchRat]*[@AC],0)"
PushS O, "ZHT0Amt =IF([@Taxed]=""Y"",[@RatZHT0]*[@AC],0)"

'Dif
PushS O, "RatDif  =Round([@BchRat]-[@RatZHT0],0)"
PushS O, "AmtDif  =Round(IF([@Taxed]=""Y"",[@BchAmt]-[@ZHT0Amt],0),0)"
FmllnyWsDta = O
End Function


Private Sub Class_Initialize()
Set WPm = MH.MB52Pm
End Sub
