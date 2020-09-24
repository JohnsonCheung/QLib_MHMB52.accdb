Attribute VB_Name = "MxFtp_A_FtpChk"
'Chk:FunSubSfx The FunSub will throw if fail
Option Compare Text
Option Explicit
Const CMod$ = "MxFtp_A_FtpChk."
Type FtpPm: Adr As String: UsrId As String: Pwd As String: End Type
Function sampFtpPm() As FtpPm
With sampFtpPm
    .UsrId = "USER"
    .Pwd = "user1234"
End With
End Function
Private Sub B_ChkFtpFdrNExi()
Dim P As FtpPm
P.Adr = "127.0.0.1"
P.UsrId = "user"
P.Pwd = "user1234"
ChkFtpFdrNExi P, "AAA"
End Sub
Sub ChkFtpFdrNExi(P As FtpPm, Fdr$)
WChkLogin P
If Not WHasFtpFdr(P, Fdr) Then Exit Sub
Ftp_Brw P.Adr
Raise "There is already folder[" & Fdr & "] in ftp address[" & P.Adr & "].  Please remove the folder!"
End Sub
Private Function WHasFtpFdr(P As FtpPm, Fdr$) As Boolean
Dim Pth$: Pth = PthTmpInst("FtpChk_FdrExit")
Dim FcmdCrt$: FcmdCrt = Pth & "HasFtpFdr.Cmd"
Dim FfnStdout$: FfnStdout = Pth & "HasFtpFdr.Stdout.txt"
Dim StdinFfn$: StdinFfn = Pth & "HasFtpFdr.stdin.txt"
WrtStr Ftp_CmdCxt(StdinFfn, FfnStdout), FcmdCrt
WrtStr WFtpScriptOfLoginAndCd(P, Fdr), StdinFfn
Shell FcmdCrt, vbHide
Dim J%: For J = 1 To 10
    Wait
    If HasFfn(FfnStdout) Then WHasFtpFdr = WHasFtpFdr_ByFfnStdOut(FfnStdout): Exit Function
Next
Raise Ftp_Msg("Ftp check folder timeout (10 seconds)", P)
End Function
Private Function WHasFtpFdr_ByFfnStdOut(FfnStdout$) As Boolean
Dim StdoutCxt$: StdoutCxt = LinesFt(FfnStdout)
Dim A$(): A = SplitCrLf(StdoutCxt)
Dim L: For Each L In Itr(A)
    If L = "250 CWD command successful." Then WHasFtpFdr_ByFfnStdOut = True: Exit Function
Next
End Function
Private Function WFtpScriptOfLoginAndCd(P As FtpPm, Fdr$) As String()
WFtpScriptOfLoginAndCd = Sy( _
    "Open " & P.Adr, _
    P.UsrId, _
    P.Pwd, _
    "cd " & Fdr, _
    "quit")
End Function
Private Sub B_WChkLogin()
Dim P As FtpPm
P.Adr = "127.0.0.1"
P.UsrId = "user"
P.Pwd = "user12341"
WChkLogin P
End Sub
Sub WChkLogin(P As FtpPm)
Dim Er$():
    Dim Pth$: Pth = PthTmpInst("FtpLogin")
    Dim FcmdCrt$: FcmdCrt = Pth & "LoginFtp.Cmd"
    Dim FfnStdout$: FfnStdout = Pth & "LoginFtp.Stdout.txt"
    Dim FfnScript$: FfnScript = Pth & "LoginFtp.script.txt"
    WrtStr Ftp_CmdCxt(FfnScript, FfnStdout), FcmdCrt
    WrtStr WFtpScriptOfLogin(P), FfnScript
    Shell FcmdCrt, vbHide
    Dim J%: For J = 1 To 5
        Wait
        If HasFfn(FfnStdout) Then Er = WErOfLogin(FfnStdout, P): GoTo Nxt
    Next
Nxt:
Er = Sy(Ftp_Msg("Login Ftp timeout (5 seconds)", P))
ChkEry Er, "LoginFtp", "Cannot login Ftp"
End Sub

Private Function WFtpScriptOfLogin(P As FtpPm) As String() ' Context of FtpLoginScript file
WFtpScriptOfLogin = Sy( _
    "Open " & P.Adr, _
    P.UsrId, _
    P.Pwd, _
    "close", _
    "quit")
End Function
Private Function WErOfLogin(FfnStdout$, P As FtpPm) As String()
Dim L$(): L = LyFt(FfnStdout)
If WIsLoginOk(L) Then Exit Function
WErOfLogin = Sy(Ftp_Msg("Login Ftp error", P))
End Function
Private Function WIsLoginOk(StdoutCxt$())
Dim L: For Each L In Itr(StdoutCxt)
    If L = "230 User logged in." Then WIsLoginOk = True: Exit Function
Next
End Function
