Attribute VB_Name = "MxIde_Dv_Enm"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Deri_Enm."

Function MdnyNeedDerienm() As String()
Dim O$()
Dim C As VBComponent: For Each C In CPj.VBComponents
    With SrcoptDvenmzCmp(C)
        If .Som Then PushI MdnyNeedDerienm, C.Name
    End With
Next
End Function
