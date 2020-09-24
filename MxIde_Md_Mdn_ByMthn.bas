Attribute VB_Name = "MxIde_Md_Mdn_ByMthn"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Md_Mdn_ByMthn."

Function MdnyMthnPC(Mthn) As String(): MdnyMthnPC = MdnyMthnP(CPj, Mthn): End Function
Function MdnyMthnP(P As VBProject, Mthn) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    If HasMthnM(C.CodeModule, Mthn) Then PushI MdnyMthnP, C.Name
Next
End Function
