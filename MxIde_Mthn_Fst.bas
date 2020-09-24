Attribute VB_Name = "MxIde_Mthn_Fst"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthn_Fst."

Function MthnFstM$(M As CodeModule):  MthnFstM = MthnL(BixMthFst(SrcM(M))): End Function
Function MthnFstMC$():               MthnFstMC = MthnFstM(CMd):             End Function
