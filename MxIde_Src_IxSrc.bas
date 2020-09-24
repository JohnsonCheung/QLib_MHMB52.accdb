Attribute VB_Name = "MxIde_Src_IxSrc"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_Ix."

Function IxSrcFst%(Src$()): IxSrcFst = IxSrcNxt(Src, 0): End Function
Function IxSrcNxt%(Src$(), Bix%)
Dim O%: For O = Bix + 1 To UB(Src)
    If IsLnCd(Src(O)) Then IxSrcNxt = O: Exit Function
Next
IxSrcNxt = -1
End Function
