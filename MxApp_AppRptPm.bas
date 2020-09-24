Attribute VB_Name = "MxApp_AppRptPm"
Option Compare Text
Option Explicit
Const CMod$ = "MxApp_AppRptPm."
Type Oup
    A As String
End Type
Type Inp
    A As String
End Type
Type RptPm
    Inp As Inp
    Oup As Oup
End Type

Function RptPmTp$()
'Fx: Fxn Fx
'    MB52 C
'    MXAA C:\AsX_
'FxCol: Fxn Wsn Fldn ColTy Coln
'    MB52 Sheet1
'       Fldn TorN ASD
'Ws: Fbn Wsnn
'    MB52 Sheet Sheet2
'    MXAA Sheet1 SHeet2
'Fb: Fbn Fb
'    DDD C:\sdfsdf
'    Ff  C:\sdfdf
'Fbt: Fbn Tbnn
'    DDD AA BB
'    CCC BB DD
'Fxw: Fxn Wsnn
'    AAA MB52 Sheet1
'    BBB MB52 Sheet1
'Fxo: Fxon Fxo
'    MB32 C:\LJKLKJDf
'Fxow: Fxon Wsnn
'Fxi: Fx
'OupPt
'
Const A_1$ = "E Mem | Mem Req AlZZLen" & _
vbCrLf & "E Txt | Txt Req" & _
vbCrLf & "E Crt | Dte Req Dft=Now" & _
vbCrLf & "E Dte | Dte" & _
vbCrLf & "F Amt * | *Amt" & _
vbCrLf & "F Crt * | CrtDte" & _
vbCrLf & "F Dte * | *Dte" & _
vbCrLf & "F Txt * | Fun * Txt" & _
vbCrLf & "F Mem * | Lines" & _
vbCrLf & "T Sess | * CrtDte" & _
vbCrLf & "T Msg  | * Fun *Txt | CrtDte" & _
vbCrLf & "T Lg   | * Sess Msg CrtDte" & _
vbCrLf & "T LgV  | * Lg Lines" & _
vbCrLf & "D . Fun | Function name that call the log" & _
vbCrLf & "D . Fun | Function name that call the log" & _
vbCrLf & "D . Msg | it will a new record when Lg-function is first time using the Fun+MsgTxt" & _
vbCrLf & "D . Msg | ..."
'LnkSpecTp = A_1
End Function
