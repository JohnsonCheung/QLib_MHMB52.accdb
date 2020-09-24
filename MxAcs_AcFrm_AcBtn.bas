Attribute VB_Name = "MxAcs_AcFrm_AcBtn"
Option Compare Text
Option Explicit
Const CMod$ = "MxAcs_AcFrm_AcBtn."
Const X_MsgExc$ = "EXCESS *_Click()"
Const X_MsgMis$ = "MISSING *_Click()"
Sub SetAcBtnEvtProc()
Dim F: For Each F In FrmnyC
    Debug.Print F; "......"
    DoCmd.OpenForm F, acDesign, , , , acHidden
    WSetBtn Forms(F)
    DoCmd.Close acForm, F, acSaveYes
Next
End Sub
Private Sub WSetBtn(F As Access.Form)
Dim C As Access.Control: For Each C In F.Controls
    If TypeName(C) = "CommandButton" Then
        With CvBtn(C)
            If .OnClick = "" Then
                .OnClick = "[Event Procedure]"
                Debug.Print F.Name, C.Name
            End If
        End With
    End If
Next
End Sub
Sub RenAcBtn(Frmn$)
Dim O$()
Dim N: For Each N In AcBtnny(FrmFrmn(Frmn))
    PushI O, FmtQQ("Forms(""?"").Controls(""?"").Name = ""?""", Frmn, N, N)
Next
D FmtT1ry(SySrtQ(O), NoIx:=True)
End Sub

Sub ChkAcBtn(Frmn$)
Dim NySub$(): NySub = W2NySub(Frmn)
Dim NyFrm$(): NyFrm = WBtnMthnyzFrmn(Frmn)
Dim NyExc$(): NyExc = SySrtQ(SyMinus(NySub, NyFrm))
Dim NyMis$(): NyMis = SySrtQ(SyMinus(NyFrm, NySub))
ChkEry SyAdd(MsgHdrAy(X_MsgExc, NyExc), MsgHdrAy(X_MsgMis, NyMis))
End Sub

Private Function W2NySub(Frmn$) As String()
DoCmd.OpenForm Frmn, acDesign, , , , acHidden
Dim Frm As Access.Form: Set Frm = Forms(Frmn)
If Frm.HasModule Then
    Dim N$()
        N = Mthny(SrcCmp(Cmp("Form_" & Frmn)))
        N = AmRmvSfx(AwSfx(N, "_Click"), "_Click")
    PushIAy W2NySub, AmAddPfx(N, "Form_" & Frmn & ".")
End If
End Function
Private Sub B_WBtnMthnyzFrmn(): Dmp WBtnMthnyzFrmn("Rpt"): End Sub
Private Function WBtnMthnyzFrmn(Frmn$) As String()
DoCmd.OpenForm Frmn, acDesign, , , , acHidden
PushIAy WBtnMthnyzFrmn, AmAddPfx(AcBtnny(Forms(Frmn)), "Form_" & Frmn & ".")
DoCmd.Close acForm, Frmn, acSaveNo
End Function

Sub ChkAcBtnAllFrm()
Dim NySub$(): NySub = WBtnMthnyC
Dim NyFrm$(): NyFrm = W3NyFrm
Dim NyExc$(): NyExc = SySrtQ(SyMinus(NySub, NyFrm))
Dim NyMis$(): NyMis = SySrtQ(SyMinus(NyFrm, NySub))
ChkEry SyAdd(MsgHdrAy(X_MsgExc, NyExc), MsgHdrAy(X_MsgMis, NyMis))
End Sub
Private Function WBtnMthnyC() As String()
Dim F: For Each F In FrmnyC
    DoCmd.OpenForm F, acDesign, , , , acHidden
    Dim Frm As Access.Form: Set Frm = Forms(F)
    If Frm.HasModule Then
        Dim N$()
            N = Mthny(SrcCmp(Cmp("Form_" & F)))
            N = AmRmvSfx(AwSfx(N, "_Click"), "_Click")
        PushIAy WBtnMthnyC, AmAddPfx(N, F & " ")
    End If
Next

End Function
Private Sub B_W3NyFrm(): Dmp W3NyFrm: End Sub
Private Function W3NyFrm() As String()
Dim F: For Each F In FrmnyC
    DoCmd.OpenForm F, acDesign, , , , acHidden
    PushIAy W3NyFrm, AmAddPfx(AcBtnny(Forms(F)), F & " ")
    DoCmd.Close acForm, F, acSaveNo
Next
End Function

Sub DmpAcBtnMisEvtProc(): Dmp W4Ly_5: End Sub

Private Function W4Ly_5() As String()
Dim F: For Each F In AeEle(FrmnyC, "Switchboard")
    DoCmd.OpenForm F, acDesign, , , , acHidden
    Dim C As Access.Control: For Each C In Forms(F).Controls
        PushNB W4Ly_5, WLnFrmn(C, F)
    Next
    DoCmd.Close acForm, F, acSaveNo
Next
End Function
Private Function WLnFrmn$(C As Access.Control, Frmn)
If TypeName(C) <> "CommandButton" Then Exit Function
Dim M$
    M = CvBtn(C).OnClick
If M = "[Event Procedure]" Then Exit Function
WLnFrmn = Frmn & " " & C.Name & " " & M
End Function
