Attribute VB_Name = "MxIde_Md_Op_EndTrimMd"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_SrcOp_EndTrim."
Sub EndtrimMdM(M As CodeModule)
Const CSub$ = CMod & "EndTrimMdM"
Dim J%, Trimmed As Boolean
While WHasLnEndBlnk(M)
    ThwLoopTooMuch CSub, J
    M.DeleteLines M.CountOfLines, 1
    Trimmed = True
Wend
If Trimmed Then Debug.Print "EndTrimMd: Module is trimmed [" & Mdn(M) & "]"
End Sub
Private Function WHasLnEndBlnk(M As CodeModule) As Boolean
Dim N&: N = M.CountOfLines
    If N = 0 Then Exit Function
    If IsLnCd(M.Lines(N, 1)) Then Exit Function
WHasLnEndBlnk = True
End Function

Sub EndtrimMdPC(): W2EndTrimMdP CPj: End Sub
Private Sub W2EndTrimMdP(P As VBProject)
Dim C As VBComponent: For Each C In P.VBComponents
    EndtrimMdM C.CodeModule
Next
End Sub

