Attribute VB_Name = "MxXls_Rg_RgPrp"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_A_Rg_XOp."
Function LnoLas%(R As Range):           LnoLas = R.Column + NColRg(R) - 1: End Function
Function RnoLas&(R As Range):           RnoLas = R.Row + NRowRg(R) - 1:    End Function
Function SyRgr(R As Range) As String():  SyRgr = SyRgr(RgR(R)):            End Function
Function SyRgc(R As Range) As String():  SyRgc = SySqc(RgC(R).Value):      End Function
