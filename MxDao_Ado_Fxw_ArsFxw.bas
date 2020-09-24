Attribute VB_Name = "MxDao_Ado_Fxw_ArsFxw"
Option Compare Text
Option Explicit

Function ArsFxwc(Fx$, W$, C, Optional Bepr$) As ADODB.Recordset: Set ArsFxwc = ArsCnq(CnFx(Fx), SqlSelFld(Axtn(W), C, Bepr)): End Function
Function ArsFxwcDis(Fx$, W$, disColn$, Optional Bepr$) As ADODB.Recordset: Set ArsFxwcDis = ArsCnq(CnFx(Fx), SqlSelFf(Axtn(W), disColn, IsDis:=True, Bepr:=Bepr)): End Function
