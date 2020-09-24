Attribute VB_Name = "MxDta_Da_Dt"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Dt."

Sub DmpDt(A As Dt):                             DmpAy FmtDt(A):                       End Sub
Function DtDrpDc(D As Dt, CC$) As Dt: DtDrpDc = DtDrs(DrsDrpDc(DrsDt(D), CC), D.Dtn): End Function
Function DtFf(Dtn$, FF$, Dy()) As Dt:    DtFf = Dt(Dtn, SplitSpc(FF), Dy):            End Function
Function DtyEmp() As Dt(): End Function
Function IsEmpDt(D As Dt) As Boolean: IsEmpDt = Si(D.Dy) = 0: End Function

Function DrsDrpDc(D As Drs, CC$) As Drs
Dim CiyDrp%(): CiyDrp = InyDrsCc(D, CC)
Dim FnyNw$(): FnyNw = FnyMinusFF(D.Fny, CC)
Dim DyNw(): DyNw = DyDrpDc(D.Dy, CiyDrp)
DrsDrpDc = Drs(FnyNw, DyNw)
End Function

Function DyDrpDc(Dy(), Ciy%()) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    PushI DyDrpDc, AeIxy(Dr, Ciy)
Next
End Function
