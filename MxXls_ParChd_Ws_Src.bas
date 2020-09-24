Attribute VB_Name = "MxXls_ParChd_Ws_Src"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_ParChd_Ws_Src."
Sub Worksheet_SelectionChange(ByVal Target As Range)
PutLoParChd Target
End Sub
