Attribute VB_Name = "MxIde_Mth_Slm_zTool_CprSlmb"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_Slmb_SlmCpr."
Sub CprSlmbPC():                                  CprSlmb MdnyPC:          End Sub
Sub CprSlmbInlAliPC():                            CprSlmb MdnyPC, True:    End Sub
Sub CprSlmbMdn(Mdn$, Optional InlAli As Boolean): CprSlmb Sy(Mdn), InlAli: End Sub
Sub CprSlmbMC(Optional InlAli As Boolean):        CprSlmb CMdny, InlAli:   End Sub
Private Sub CprSlmb(Mdny$(), Optional InlAli As Boolean)
Dim Cpmdy() As Cpmd
Dim N: For Each N In Itr(Mdny)
    PushCpmdy Cpmdy, CpmdyMdn(N, InlAli)
Next
CprCpmd Cpmdy
End Sub
Private Function CpmdyMdn(Mdn, InlAli As Boolean) As Cpmd()
Dim Slmb: For Each Slmb In Itr(SlmbySrc(SrcMdn(Mdn)))
    With SlmboptAli(CvSy(Slmb))
        If .Som Then
            PushCpmd CpmdyMdn, Cpmd(Mdn, Slmb, .Ly)
        Else
            If InlAli Then
                PushCpmd CpmdyMdn, Cpmd(Mdn, Slmb, Slmb)
            End If
        End If
    End With
Next
End Function
