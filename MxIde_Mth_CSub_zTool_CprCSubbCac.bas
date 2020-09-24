Attribute VB_Name = "MxIde_Mth_CSub_zTool_CprCSubbCac"
Option Compare Database
Option Explicit
Sub CprCSubbFtcacPC()
EnsFtcacPC
CprCSubbFtcac MdnyFtcacMthYesP(CPj)
End Sub
Sub CprCSubbFtcacMdn(Mdn$, Optional InlEns As Boolean): CprCSubbFtcac Sy(Mdn), InlEns: End Sub
Sub CprCSubbFtcacMC(Optional InlEns As Boolean):        CprCSubbFtcac CMdny, InlEns:   End Sub
Private Sub CprCSubbFtcac(Mdny$(), Optional InlEns As Boolean)
Const CSub$ = ""
Dim Cpmdy() As Cpmd
Dim J%
Debug.Print "CprCSubbFtcac: NMd="; Si(Mdny)
Dim N: For Each N In Itr(Mdny)
    J = J + 1: If J Mod 10 = 0 Then Debug.Print J;: If J Mod 100 = 0 Then Debug.Print
    PushCpmdy Cpmdy, CpmdyMthyy(N, MthyyFtcacMdn(N), InlEns)
Next
CprCpmd Cpmdy
End Sub
Private Function CpmdyMthyy(Mdn, Mthyy(), InlEns As Boolean) As Cpmd()
'>Cpmdy is comparing the all @mthyy if they are diff and optionally (@InlEns) include those are no dif but UseCSub.
'One compared mth will be one :Cpmd element.
Dim I: For Each I In Itr(Mthyy)
    Dim Mthy$(): Mthy = I
    Dim IsUse As Boolean, Has As Boolean
    IsUse = IsUseCSub(Mthy)
    Has = HasCSub(Mthy)
    If IsUse Or Has Then
        With CSubboptEns(Mthy)
            If .Som Then
                PushCpmd CpmdyMthyy, Cpmd(Mdn, Mthy, .Ly)
            Else
                If InlEns Then
                    If IsUse Then
                        PushCpmd CpmdyMthyy, Cpmd(Mdn, Mthy, Mthy)
                    End If
                End If
            End If
        End With
    End If
Next
End Function
Private Function HasCSub(Mthy$()) As Boolean
Dim J%: For J = 1 To UB(Mthy) - 1
    If IsLnCSub(Mthy(J)) Then HasCSub = True: Exit Function
Next
End Function
