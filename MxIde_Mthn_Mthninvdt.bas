Attribute VB_Name = "MxIde_Mthn_Mthninvdt"
Option Compare Text
Const CMod$ = "MxIde_Mthn_Mthninvdt."
Option Explicit
Private Sub B_X1_Is_PrvMthn_InVdt()
GoSub T1
Exit Sub
Dim MthnPrv$
T1:  MthnPrv = "X_AA": Ept = False: GoTo Tst
T2:  MthnPrv = "W1_AA": Ept = False: GoTo Tst
T3:  MthnPrv = "XW_AA": Ept = True: GoTo Tst
Tst:
    Act = X1_Is_PrvMthn_Invdt(MthnPrv)
    C
    Return
End Sub
Sub MthnInvdtTop10():                               Dmp X_Ly_ForMdny1(MdnyPC, 10):             End Sub
Sub MthnPrvInvdtTop10(Optional Wh_NMth_LE As Byte): Dmp X_Ly_ForMdny1(MdnyPC, 10, Wh_NMth_LE): End Sub
Sub MthnPrvInvdtAll():                              Vc X_Ly_ForMdny1(MdnyPC):                  End Sub
Sub MthnPrvInvdtMC():                               Dmp X_Ly_ForMdny1(Sy(CMdn)):               End Sub
Private Function X_Ly_ForMdny1(Mdny$(), Optional Top% = 32767, Optional Wh_NMth_LE As Byte) As String()
Dim IxMd%: IxMd = 1
Dim M: For Each M In Itr(SySrtQ(Mdny))
    PushIAy X_Ly_ForMdny1, X1_Ly_ForMdn(M, Wh_NMth_LE, IxMd)
    If IxMd > Top Then Exit Function
Next
End Function
Private Function X1_Ly_ForMdn(Mdn, Wh_NMth_LE As Byte, OIx%) As String()
Dim Ny$(): Ny = X1_Mthny_ofPrv_ofInvdt(Mdn, Wh_NMth_LE): If Si(Ny) = 0 Then Exit Function
X1_Ly_ForMdn = SyStrSy(OIx & " " & Mdn, LyTab(Ny))
OIx = OIx + 1
End Function
Private Sub B_X1_Mthny_ofPrv_ofInvdt()
GoSub T1
Exit Sub
Dim Mdn$
T1:
    Mdn = CMdn
    GoTo Tst
Tst:
    Act = X1_Mthny_ofPrv_ofInvdt(Mdn, 0)
    Dmp Act
    Return

End Sub
Private Function X1_Mthny_ofPrv_ofInvdt(Mdn, Wh_NMth_LE As Byte) As String()
Dim NyPrv$(): NyPrv = MthnyPrvM(Md(Mdn))
Dim N: For Each N In Itr(NyPrv)
    If X1_Is_PrvMthn_Invdt(N) Then
        Push X1_Mthny_ofPrv_ofInvdt, N
        If Wh_NMth_LE > 0 Then
            Dim J As Byte: J = J + 1
            If J > Wh_NMth_LE Then Exit Function
        End If
    End If
Next
End Function
Private Function X1_Is_PrvMthn_Invdt(MthnPrv) As Boolean
Static PfxyOk1$(), X1 As Boolean: If Not X1 Then X1 = True: PfxyOk1 = SplitSpc("B_ Cmd_ X_ W_ Z_ ZZ_") 'ZZ is Tool
Select Case True
Case _
    HasPfxy(MthnPrv, PfxyOk1)
    Exit Function
End Select

Select Case ChrFst(MthnPrv)
Case "X", "W", "Z"
Case Else: Exit Function
End Select

Dim S4$: S4 = Left(MthnPrv, 4)
Dim C2$: C2 = ChrSnd(S4)
Dim IsC2Dig As Boolean: IsC2Dig = IsDig(C2)
Dim IsC3Dig As Boolean: IsC3Dig = IsDig(ChrThd(S4))

Select Case True
Case _
    IsC2Dig And IsC3Dig And Mid(S4, 4, 1) = "_", _
    IsC2Dig And Mid(S4, 3, 1) = "_"
Case Else
    X1_Is_PrvMthn_Invdt = True
End Select
End Function

