Attribute VB_Name = "MxIde_Dcl_Drs"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dcl_Drs."
Public Const FFDcl$ = "Pjn Mdn Dcll"

Function DrsTDclPC() As Drs: DrsTDclPC = DrsTDclP(CPj): End Function
Function DrsTDclP(P As VBProject) As Drs
Dim Dy(), Pjn$
Pjn = P.Name
Dim C As VBComponent: For Each C In P.VBComponents
    PushI Dy, Array(Pjn, C.Name, DcllCmp(C))
Next
DrsTDclP = DrsFf(FFDcl, Dy)
End Function
