Attribute VB_Name = "MxIde_Gen_BldCdy"
Option Compare Text
Option Explicit

Function CdyIf(PfxPrv$, Mthn$, IsGen As Boolean, TpyCdl$()) As String()
If Not IsGen Then Exit Function
Dim Tp: For Each Tp In TpyCdl
    PushI CdyIf, PfxPrv & FmtQQ(CStr(Tp), Mthn)
Next
End Function
