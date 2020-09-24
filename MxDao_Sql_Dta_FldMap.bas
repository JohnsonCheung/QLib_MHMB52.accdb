Attribute VB_Name = "MxDao_Sql_Dta_FldMap"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Sql_TmlFldMap."
Sub AsgTmlFldMap(TmlFldMap$, OFnyA$(), OFnyB$())
Erase OFnyA, OFnyB
Dim TmFldMap: For Each TmFldMap In ItrTml(TmlFldMap)
    With Brk1Colon(TmFldMap)
        PushI OFnyA, .S1
        PushI OFnyB, .S2
    End With
Next
End Sub
