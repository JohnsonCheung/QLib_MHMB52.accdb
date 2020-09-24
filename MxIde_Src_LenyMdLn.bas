Attribute VB_Name = "MxIde_Src_LenyMdLn"
Option Compare Text
Const CMod$ = "MxIde_Src_LenyMdLn."
Option Explicit
Private Sub B_FmtMdLen()
BrwAy FmtT1ry(FmtMdLen, CiiAliR:="0", CiiHypKey:="0-", CiiNbr:="0")
End Sub
Function FmtMdLen() As String()
Dim C As VBComponent: For Each C In CPj.VBComponents
    PushI FmtMdLen, WdtLines(SrclCmp(C)) & " " & C.Name
Next
End Function
