Attribute VB_Name = "MxDta_Da_Drs_DrsNw"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Drs_DrsNw."
Function DrsTmy4R(Ff5$, Tmly$()) As Drs
Dim I, Dy(): For Each I In Itr(Tmly)
    PushI Dy, Tmy4r(I)
Next
DrsTmy4R = DrsFf(Ff5, Dy)
End Function
Function DrsSSy(Ssy$(), FF$) As Drs: DrsSSy = DrsFf(FF, DyoSSy(Ssy)): End Function
Function DrsTRstLy(T1ry$(), FF$) As Drs
Dim I, Dy(): For Each I In Itr(T1ry)
    PushI Dy, Tmy2r(I)
Next
DrsTRstLy = DrsFf(FF, Dy)
End Function

Function DrsEmpFf(FF$) As Drs:         DrsEmpFf = DrsFf(FF, AvEmp):   End Function
Function DrsFf(FF$, Dy()) As Drs:         DrsFf = Drs(FnyFF(FF), Dy): End Function
Function DrsDy(D As Drs, Dy()) As Drs:    DrsDy = Drs(D.Fny, Dy):     End Function
