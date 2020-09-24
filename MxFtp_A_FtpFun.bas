Attribute VB_Name = "MxFtp_A_FtpFun"
Option Compare Text
Option Explicit
Const CMod$ = "MxFtp_A_FtpFun."
Function Ftp_CmdCxt$(FfnScript$, FfnStdout$)
'Const Pause$ = vbCrLf & "pause"
Const Pause$ = ""
Ftp_CmdCxt = FmtQQ("ftp -s:""?"" > ""?""", FfnScript, FfnStdout) & Pause ' Context of FcmdCrt to login Ftp site
End Function
Sub Ftp_Brw(Optional FtpAdr$ = "127.0.0.1", Optional Fdr$): Shell FmtStr("explorer.exe ""ftp://{0}/{1}""", FtpAdr, Fdr), vbMaximizedFocus: End Sub
Function Ftp_Msg$(Tit$, P As FtpPm)
Dim O$(): O = Box(Tit)
PushI O, "Ftp address : [" & P.Adr & "]"
PushI O, "User id     : [" & P.UsrId & "]"
PushI O, "Password    : [******]"
PushI O, "======================"
Ftp_Msg = JnCrLf(O) & vbCrLf
End Function
