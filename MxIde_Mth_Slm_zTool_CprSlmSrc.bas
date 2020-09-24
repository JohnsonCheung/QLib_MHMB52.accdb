Attribute VB_Name = "MxIde_Mth_Slm_zTool_CprSlmSrc"
Option Compare Text
Option Explicit

Sub slmzCprSrcPC():      slmzCprSrcMdny MdnySrcPC: End Sub
Sub slmzCprSrcMdn(Mdn$): slmzCprSrcMdny Sy(Mdn):   End Sub
Sub slmzCprSrcMC():      slmzCprSrcMdny CMdny:     End Sub
Private Sub slmzCprSrcMdny(Mdny$())
Dim O() As Cpmd
    Dim M() As Nly: M = MsrcySlmMdny(Mdny)
    Dim J&: For J = 0 To UbNly(M)
        With M(J)
            PushCpmd O, Cpmd(.Nm, SrcMdn(.Nm), .Ly)
        End With
    Next
CprCpmd O
End Sub
