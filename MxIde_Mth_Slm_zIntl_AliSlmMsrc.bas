Attribute VB_Name = "MxIde_Mth_Slm_zIntl_AliSlmMsrc"
Option Compare Database
Option Explicit

Function MsrcySlmMdny(Mdny$()) As Nly()
Dim N: For Each N In Itr(Mdny)
    With SrcoptSlm(SrcMdn(N))
        If .Som Then
            PushNly MsrcySlmMdny, Nly(N, .Ly)
        End If
    End With
Next
End Function
Function MsrcySlmPC() As Nly():                      MsrcySlmPC = MsrcySlmP(CPj):            End Function
Private Function MsrcySlmP(P As VBProject) As Nly():  MsrcySlmP = MsrcySlmMdny(MdnySrcP(P)): End Function
