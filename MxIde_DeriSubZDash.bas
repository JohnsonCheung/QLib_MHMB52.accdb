Attribute VB_Name = "MxIde_DeriSubZDash"
Option Compare Text
Const CMod$ = "MxIde_DeriSubZDash."
Option Explicit

Sub DeriSubZDashMC()

End Sub
Sub DeriSubZDashM(M As CodeModule)

End Sub
Sub DeriSubZDashPC()
Dim C As VBComponent: For Each C In CPj.VBComponents
    DeriSubZDashM MdCmp(C)
Next
End Sub
