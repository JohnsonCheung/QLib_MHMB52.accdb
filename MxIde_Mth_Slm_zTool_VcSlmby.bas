Attribute VB_Name = "MxIde_Mth_Slm_zTool_VcSlmby"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_Slmb_SlmBrw."
Private Sub B_slmzVcSlmbyMdn():            slmzVcSlmbyMdn "MxIde_Mth_Slm_zTool_BrwSlmby": End Sub
Sub slmzVcSlmbyPC():                       slmzVcSlmbyP CPj:                              End Sub
Sub slmzVcSlmbyMC():                       slmzVcSlmbyM CMd:                              End Sub
Sub slmzVcSlmbyMdn(Mdn$):                  slmzVcSlmbyM MdMdn(Mdn):                       End Sub
Private Sub slmzVcSlmbyM(M As CodeModule): BrwLyy SlmboyM(M):                             End Sub
Private Sub slmzVcSlmbyP(P As VBProject):  BrwLyy SlmboyP(P):                             End Sub
Private Function SlmboyP(P As VBProject) As Variant()
Dim Mdno%: Mdno = 1
Dim C As VBComponent: For Each C In P.VBComponents
    Dim Blky(): Blky = SlmboyM(C.CodeModule, Mdno)
    If Si(Blky) > 0 Then
        PushIAy SlmboyP, Blky
        Mdno = Mdno + 1
    End If
Next
Stop
End Function
Private Function SlmboyM(M As CodeModule, Optional Mdno% = 1) As Variant()
Dim SLmby(): SLmby = SlmbySrc(SrcM(M))
Dim N$: N = Mdn(M)
Dim ISlmb: For Each ISlmb In Itr(SLmby)
    PushI SlmboyM, Slmbo(CvSy(ISlmb), Mdno, N)
Next
End Function
Private Function Slmbo(Slmb$(), Mdnno%, Mdn$) As String()   '#Slmbo:Slm-Blk-Oup# Slm blk oup for Brw
Dim Hdr$(): Hdr = LyUL("Md#(" & Mdnno & ") " & Mdn)
Slmbo = SyAdd(Hdr, Slmb)
End Function
