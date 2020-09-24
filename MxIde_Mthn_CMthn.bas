Attribute VB_Name = "MxIde_Mthn_CMthn"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthn_CMthn."

Function CMthnM$(M As CodeModule)
Dim K As vbext_ProcKind
CMthnM = M.ProcOfLine(CLnoM(M), K)
End Function

Function CMthn$(): CMthn = CMthnM(CMd): End Function
