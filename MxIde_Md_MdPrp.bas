Attribute VB_Name = "MxIde_Md_MdPrp"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Md_Prp."
Function MdnPfx$(PfxMdn$, Optional C As eCas):    MdnPfx = ElePfx(SySrtQ(MdnyPC), PfxMdn): End Function
Function CMdn():                                    CMdn = CCmp.Name:                      End Function
Function Mdn(M As CodeModule):                       Mdn = M.Parent.Name:                  End Function
Function CMd() As CodeModule:                    Set CMd = CPne.CodeModule:                End Function
Function CMdnDot$():                             CMdnDot = MdnDotM(CMd):                   End Function
Function MdnyDotPC() As String():              MdnyDotPC = MdnyDotP(CPj):                  End Function
Function MdnDotM$(M As CodeModule):              MdnDotM = PjnM(M) & "." & Mdn(M):         End Function
Function MdnyDotP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushI MdnyDotP, MdnDotM(C.CodeModule)
Next
End Function
