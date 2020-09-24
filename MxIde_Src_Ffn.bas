Attribute VB_Name = "MxIde_Src_Ffn"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_Ffn."

Function FfnSrcMdn$(Mdn$, Pth$):          FfnSrcMdn = FfnSrcM(MdMdn(Mdn), Pth): End Function
Function FfnSrcM$(M As CodeModule, Pth$):   FfnSrcM = FtSrc(M.Parent, Pth):     End Function
Function FfnSrcMC$():                      FfnSrcMC = FtSrc(CCmp, PthPC):       End Function

Private Sub B_FtySrcPC():                        Vc FtySrcPC:           End Sub
Function FtSrc$(C As VBComponent, Pth$): FtSrc = PthSrcCmp(C) & WFn(C): End Function
Private Function WFn$(C As VBComponent):   WFn = C.Name & WExt(C.Type): End Function
Private Function WExt$(A As vbext_ComponentType)
Const CSub$ = CMod & "WExt"
Dim O$
Select Case A
Case vbext_ct_ClassModule: O = ".cls"
Case vbext_ct_StdModule: O = ".bas"
Case vbext_ct_Document: Thw CSub, "Unexpected CmpTy=Document.  To get extension of Form/Report should use Acs.ObjTy"
Case Else: Thw CSub, "Unexpected CmpTy.  Should be [Class | Module].  The [Document]-CmpTy(Which is Report|Form should must use Access.Application.SaveAsText(which is hidden)"
End Select
WExt = O
End Function

Function FtySrcPC() As String(): FtySrcPC = FtySrcP(CPj): End Function
Function FtySrcP(P As VBProject) As String()
Dim Pth$: Pth = PthSrcPC
Dim C As VBComponent: For Each C In P.VBComponents
    Select Case C.Type
    Case vbext_ct_StdModule, vbext_ct_ClassModule
        PushI FtySrcP, FtSrc(C, Pth)
    End Select
Next
End Function
