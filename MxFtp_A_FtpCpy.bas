Attribute VB_Name = "MxFtp_A_FtpCpy"
Option Compare Text
Option Explicit
Const CMod$ = "MxFtp_A_FtpCpy."
Sub CpyFfnyToFtp(Ffny$(), P As FtpPm, FdrAtFtp$)
ChkFtpFdrNExi P, FdrAtFtp
Dim Pth$: Pth = PthTmpInst("FtpCpy")
Dim FcmdCrt$: FcmdCrt = Pth & "FtpCpy.Cmd"
Dim FfnStdout$: FfnStdout = Pth & "FtpCpy.Stdout.txt"
Dim FfnScript$: FfnScript = Pth & "FtpCpy.script.txt"
WrtStr Ftp_CmdCxt(FfnScript, FfnStdout), FcmdCrt
WrtStr WFtpScriptOfLoginAndCpyTo(Ffny, FdrAtFtp, P), FfnScript
Shell FcmdCrt, vbNormalFocus
Dim J%: For J = 1 To 60 * 10
    Wait
    If HasFfn(FfnStdout) Then WChkStdoutFfn FfnStdout, P: Exit Sub
Next
Raise Ftp_Msg("Copying statments to ftp timeout (10 minutes)", P)
End Sub
Private Function WFtpScriptOfLoginAndCpyTo$(Ffny$(), FdrAtFtp$, P As FtpPm) ' Context of FtpLoginScript file
Dim O$()
PushI O, "Open " & P.Adr
PushI O, P.UsrId
PushI O, P.Pwd
PushI O, "mkdir """ & FdrAtFtp & """"
PushI O, "cd """ & FdrAtFtp & """"
Dim F: For Each F In Itr(Ffny)
PushI O, "put """ & F & """"
Next
PushI O, "close"
PushI O, "quit"
WFtpScriptOfLoginAndCpyTo = JnCrLf(O)
End Function
Private Sub WChkStdoutFfn(FfnStdout$, P As FtpPm)
BrwFt FfnStdout
MsgBox "Done", vbInformation
End Sub
