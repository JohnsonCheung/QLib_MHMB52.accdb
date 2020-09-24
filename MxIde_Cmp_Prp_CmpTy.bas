Attribute VB_Name = "MxIde_Cmp_Prp_CmpTy"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Cmp_Prp_CmpTy."

Function IsMod(C As VBComponent) As Boolean:       IsMod = C.Type = vbext_ct_StdModule:   End Function
Function IsCls(C As VBComponent) As Boolean:       IsCls = C.Type = vbext_ct_ClassModule: End Function
Function IsMd(C As VBComponent) As Boolean:         IsMd = IsMod(C) Or IsCls(C):          End Function
Function IsCmpSrc(C As VBComponent) As Boolean: IsCmpSrc = IsMd(C) Or IsCmpDoc(C):        End Function
Function IsCmpDoc(C As VBComponent) As Boolean: IsCmpDoc = C.Type = vbext_ct_Document:    End Function
