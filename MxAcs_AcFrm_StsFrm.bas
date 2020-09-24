Attribute VB_Name = "MxAcs_AcFrm_StsFrm"
Option Compare Text
Option Explicit
Const CMod$ = "MxAcs_AcFrm_StsFrm."
Private Msg$()

Private Sub B_Sts()
Dim J%: For J = 0 To 10
    Sts J
Next
Stop
End Sub
Sub StsQry(QNm$): Sts "Running query " & QNm & "....": End Sub
Sub StsLnk(T):    Sts "Linking " & T & " ........":    End Sub
Sub ShwSts():     BrwAy AyRev(Msg):                    End Sub
Sub StsDone():    Sts "Done":                          End Sub
Sub Sts(Sts)
Dim A$: A = Now & " " & Sts
PushS Msg, A
X_1RfhBtn_2
Debug.Print A
End Sub
Sub StsQQ(StsQQ$, ParamArray Ap()): Dim Av(): Av = Ap: Sts FmtQQAv(StsQQ, Av): End Sub
Sub ClrStsBtn():                    Erase Msg: X_1RfhBtn_2:                    End Sub

Private Sub X_1RfhBtn_2()
Dim B As Access.CommandButton: Set B = X2_Btn
If IsNothing(B) Then Exit Sub
B.Caption = Si(Msg) & " Msgs"
B.Requery
End Sub
Private Function X2_Btn() As Access.CommandButton
Dim F As Access.Form: Set F = CFrm: If IsNothing(F) Then Exit Function
If Not HasCtl(F, "CmdMsg") Then Exit Function
Dim C As Control: Set C = F.Controls("CmdMsg")
If TypeName(C) <> "CommandButton" Then Raise "MxSts.X2_Btn: Has a control with Name CmdMsg, but it is not a CommandButton, but[" & TypeName(C.Object) & "]"
Set X2_Btn = F.Controls("CmdMsg")
End Function

Sub ClrMainMsg()
'Assume there is Application.Forms("Main").Msg (TextBox)
'MMsg means Main.Msg (TextBox)
Dim M As TextBox: Set M = MainMsgBox
If Not IsNothing(M) Then M.Value = ""
End Sub

Sub SetMainMsgQnm(QryNm)
SetMainMsg "Running query: (" & QryNm & ")...."
End Sub

Sub SetMainMsg(Msg$)
On Error Resume Next
SetTBox MainMsgBox, Msg
End Sub

Property Get MainMsgBox() As Access.TextBox
On Error Resume Next
Set MainMsgBox = MainFrm.Controls("Msg")
End Property

Property Get MainFrm() As Access.Form
On Error Resume Next
Set MainFrm = Access.Forms("Main")
End Property

Sub SetTBox(A As Access.TextBox, Msg$)
Dim CrLf$, B$
If A.Value <> "" Then CrLf = vbCrLf
B = LinesLasN(A.Value & CrLf & Now & " " & Msg, 5)
A.Value = B
DoEvents
End Sub



Sub ClrSts(): SysCmd acSysCmdClearStatus: End Sub
