Attribute VB_Name = "MxIde_Mth_CSub_zTool_CprCSubb"
Option Compare Database
Option Explicit

Sub csubzCprBlkPC():                                  CprCSubb MdnyPC:          End Sub
Sub csubzCprBlkMdn(Mdn$, Optional InlEns As Boolean): CprCSubb Sy(Mdn), InlEns: End Sub
Sub csubzCprBlkMC(Optional InlEns As Boolean):        CprCSubb CMdny, InlEns:   End Sub
Private Sub CprCSubb(Mdny$(), Optional InlEns As Boolean)
Dim Cpmdy() As Cpmd
Dim N: For Each N In Itr(Mdny)
    PushCpmdy Cpmdy, CpmdyMdn(N, InlEns)
    Stop
Next
CprCpmd Cpmdy
End Sub
Sub AA()
Dim Mthy$()
PushI Mthy, ", CSub"
Dim A As Boolean: A = IsUseCSub(Mthy)
Stop
End Sub
Private Function CpmdyMdn(Mdn, InlEns As Boolean) As Cpmd()
Dim I: For Each I In Itr(MthyySrc(SrcMdn(Mdn)))
    Dim Mthy$(): Mthy = I
    Dim IsUse As Boolean: IsUse = IsUseCSub(Mthy)
    Dim Has As Boolean: Has = HasCSub(Mthy)
    If IsUse Or Has Then
        With CSubboptEns(Mthy)
            If .Som Then
                PushCpmd CpmdyMdn, Cpmd(Mdn, Mthy, .Ly)
            Else
                If InlEns Then
                    PushCpmd CpmdyMdn, Cpmd(Mdn, Mthy, Mthy)
                End If
            End If
        End With
    End If
Next
End Function
Private Function HasCSub(Mthy$()) As Boolean
Dim L: For Each L In Itr(Mthy)
    If IsLnCSub(L) Then HasCSub = True: Exit Function
Next
End Function
Private Function IsLnCSub(L) As Boolean: IsLnCSub = HasCnstnL(L, "CSub"): End Function
