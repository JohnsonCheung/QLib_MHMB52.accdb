Attribute VB_Name = "MxDta_Da_Op_DrpCol"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Op_DrpCol."

Function DyDrpDc(Dy(), Ciy%()) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
   Push DyDrpDc, AeIxy(Dr, Ciy)
Next
End Function

Function DrsDrpDcFny(D As Drs, Fny$()) As Drs
Dim CiyDrp%()
    CiyDrp = InyFnySub(D.Fny, Fny)
Dim FnyO$(), DyO()
    FnyO = SyMinus(D.Fny, Fny)
    DyO = DyDrpDc(D.Dy, CiyDrp)
DrsDrpDcFny = Drs(FnyO, DyO)
End Function
