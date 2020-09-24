Attribute VB_Name = "MxIde_Mthn_Mdn"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthn_Mdn."

Function MdnyPubMthPC(MthnPub) As String(): MdnyPubMthPC = MdnyPubMthP(CPj, MthnPub): End Function

Private Sub B_MdnyPubMthP()
GoSub Z
Exit Sub
Dim P As VBProject, MthnPub
Z:
    D MdnyPubMthP(CPj, "AA")
    Return
End Sub
Function MdnyPubMthP(P As VBProject, MthnPub) As String()
Dim I, M As CodeModule: For Each I In ItrModP(P)
    Set M = I
    If WHasPubMth(SrcM(M), MthnPub) Then PushI MdnyPubMthP, Mdn(M)
Next
End Function
Private Function WHasPubMth(Src$(), MthnPub_) As Boolean
Dim L: For Each L In Itr(Src)
    If MthnPub(L) = MthnPub_ Then WHasPubMth = True
Next
End Function
