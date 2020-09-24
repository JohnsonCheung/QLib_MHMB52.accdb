Attribute VB_Name = "MxDta_Da_Dr_VbTy"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Dr_VbTy."

Function TslTynDr$(Dr)
Dim O$()
Dim V: For Each V In Itr(Dr)
    PushI O, TypeName(V)
Next
TslTynDr = JnTab(O)
End Function

Function DrFst(Dy())
If Si(Dy) = 0 Then Exit Function
Dim Dr: Dr = Dy(0)
Dim J%: For J = Si(Dr) To NDcDy(Dy) - 1
    PushI Dr, ""
Next
DrFst = Dr
End Function
