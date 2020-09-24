Attribute VB_Name = "MxVb_FmtTaby"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_FmtTaby."
Function DyTsy(Tsy$()) As Variant()
Dim L: For Each L In Itr(Tsy)
    PushI DyTsy, SplitTab(L)
Next
End Function
Function TsyDy(Dy()) As String()
Dim Dr: For Each Dr In Itr(Dy)
    PushI TsyDy, JnTab(Dr)
Next
End Function
Function FmtTsy(Tsy$()) As String()
FmtTsy = FmtDy(DyTsy(Tsy))
End Function
