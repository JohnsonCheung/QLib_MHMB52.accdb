Attribute VB_Name = "MxXls_Rg_RgOp_DInsCRow"
Option Compare Text
Option Explicit

Sub InsRowAt(ByVal At As Range, Optional N = 1)
RgRREnt(At, 1, 1 + N - 1).EntireRow.Insert
' Using ByRef for @At will be changed after insert.  So use ByVal  ! <==
End Sub
Sub InsColAtEnd(ByVal At As Range, Optional N = 1)
RgCCEnt(At, 1, 1 + N - 1).EntireColumn.Insert
' Using ByRef for @At will be changed after insert.  So use ByVal  ! <==
End Sub

Sub DltRowEmp(R As Range)
Dim S As Worksheet: Set S = WsRg(R)
Dim Sq():              Sq = RCnoSq(R)
Dim Rny&():           Rny = RnyOfEmpRowFmSq(Sq)
EntRows(S, Rny).Remove
End Sub
