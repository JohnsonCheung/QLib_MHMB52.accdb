Attribute VB_Name = "MxIde_Md_Cmp"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Md_Cmp."

Function ShtCmpTyM$(M As CodeModule): ShtCmpTyM = ShtCmpTy(CmpTyM(M)): End Function
Function ShtCmpTyyMdny(P As VBProject, Mdny$()) As String()
Dim N: For Each N In Itr(Mdny)
    PushI ShtCmpTyyMdny, ShtCmpTy(P.VBComponents(N).Type)
Next
End Function

Function ShtCmpTy$(A As vbext_ComponentType)
Dim O$
Select Case A
Case vbext_ComponentType.vbext_ct_Document:    O = "Doc"
Case vbext_ComponentType.vbext_ct_ClassModule: O = "Cls"
Case vbext_ComponentType.vbext_ct_StdModule:   O = "Std"
Case vbext_ComponentType.vbext_ct_MSForm:      O = "Frm"
Case vbext_ComponentType.vbext_ct_ActiveXDesigner: O = "ActX"
Case Else: Stop
End Select
ShtCmpTy = O
End Function

Function CmpTyM(M As CodeModule) As vbext_ComponentType:            CmpTyM = M.Parent.Type:             End Function
Function CmpTyCmpn(P As VBProject, Cmpn) As vbext_ComponentType: CmpTyCmpn = P.VBComponents(Cmpn).Type: End Function
Function CmpTySht(ShtCmpTy) As vbext_ComponentType
Dim O As vbext_ComponentType
Select Case ShtCmpTy
Case "Doc": O = vbext_ComponentType.vbext_ct_Document
Case "Cls": O = vbext_ComponentType.vbext_ct_ClassModule
Case "Std": O = vbext_ComponentType.vbext_ct_StdModule
Case "Frm": O = vbext_ComponentType.vbext_ct_MSForm
Case "ActX": O = vbext_ComponentType.vbext_ct_ActiveXDesigner
Case Else: Stop
End Select
End Function
