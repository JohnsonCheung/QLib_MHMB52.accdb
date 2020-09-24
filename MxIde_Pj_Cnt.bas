Attribute VB_Name = "MxIde_Pj_Cnt"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Pj_Cnt."

Sub BrwTCmpCntVC():                                   BrwDrs DrsTCmpCntVC:                         End Sub
Function DrsTCmpCntVC() As Drs:        DrsTCmpCntVC = DrsTCmpCntV(CVbe):                           End Function
Function DrsTCmpCntV(V As VBE) As Drs:  DrsTCmpCntV = DrsFf("Pj Tot Mod Cls Doc Frm Oth", WDy(V)): End Function
Private Function WDy(V As VBE) As Variant()
Dim P As VBProject: For Each P In V.VBProjects
    PushI WDy, WDr(P)
Next
End Function
Private Function WDr(P As VBProject) As Variant()
Dim NCls%, NDoc%, NFrm%, NMod%, NOth%, NTot%
Dim C As VBComponent: For Each C In P.VBComponents
    Select Case C.Type
    Case vbext_ct_ClassModule:  NCls = NCls + 1
    Case vbext_ct_Document:     NDoc = NDoc + 1
    Case vbext_ct_MSForm:       NFrm = NFrm + 1
    Case vbext_ct_StdModule:    NMod = NMod + 1
    Case Else:                  NOth = NOth + 1
    End Select
    NTot = NTot + 1
Next
WDr = Array(P.Name, NTot, NMod, NCls, NDoc, NFrm, NOth)
End Function

Function SrcCmp(C As VBComponent) As String()
On Error Resume Next
SrcCmp = SrcM(C.CodeModule)
End Function
Function Cmp(Cmpn) As VBComponent: Set Cmp = CmpP(CPj, Cmpn): End Function
Function NClsPC%():                 NClsPC = NClsP(CPj):      End Function
Function NCmpPC%():                 NCmpPC = NCmpP(CPj):      End Function
Function NModPC%():                 NModPC = NModP(CPj):      End Function
Function NDocPC%():                 NDocPC = NDocP(CPj):      End Function
Function NOthPC%():                 NOthPC = NOthP(CPj):      End Function

Function NCmpP%(P As VBProject)
If P.Protection = vbext_pp_locked Then Exit Function
NCmpP = P.VBComponents.Count
End Function

Function NModP%(P As VBProject): NModP = NCmpTy(P, vbext_ct_StdModule):   End Function
Function NClsP%(P As VBProject): NClsP = NCmpTy(P, vbext_ct_ClassModule): End Function
Function NDocP%(P As VBProject): NDocP = NCmpTy(P, vbext_ct_Document):    End Function

Function NCmpTy%(P As VBProject, Ty As vbext_ComponentType)
If P.Protection = vbext_pp_locked Then Exit Function
Dim C As VBComponent
Dim O%
For Each C In P.VBComponents
    If C.Type = Ty Then O = O + 1
Next
NCmpTy = O
End Function

Function NOthP%(P As VBProject)
NOthP = NCmpP(P) - NClsP(P) - NModP(P) - NDocP(P)
End Function
