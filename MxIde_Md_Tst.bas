Attribute VB_Name = "MxIde_Md_Tst"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Md_Tst."

Sub CrtMdTstP(P As VBProject)
W1_CrtMd_2 P
W1_CrtTstMth P
End Sub
Private Sub W1_CrtTstMth(P As VBProject)
Stop
End Sub
Private Sub W1_CrtMd_2(P As VBProject)
Dim MdnMis$(): MdnMis = W12_MdnyMis_3(P)
Dim Mdn: For Each Mdn In MdnMis
    AddMod P, Mdn
Next
End Sub
Private Function W12_MdnyMis_3(P As VBProject) As String()
Dim NAct$(): NAct = ClsnyP(P)
Dim NEpt$(): NEpt = W13_MdnyEpt
W12_MdnyMis_3 = SyMinus(NEpt, NAct)
End Function
Private Function W13_MdnyEpt() As String()

End Function
