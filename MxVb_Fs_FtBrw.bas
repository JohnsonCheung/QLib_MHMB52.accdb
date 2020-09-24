Attribute VB_Name = "MxVb_Fs_FtBrw"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_FtBrw."
'Public Const FfnOfCodeExe$ = "C:\Users\user\AppData\Local\Programs\Microsoft VS Code\Code.exe"
'"C:\Users\user\AppData\Local\Programs\Microsoft VS Code\Code.exe" "C:\Users\user\AppData\Local\Temp\JC\N20200820_213641.txt"
Public Const vc_Pth$ = "C:\Users\user\AppData\Local\Programs\Microsoft VS Code\"
Public Const vc_binPth$ = vc_Pth & "Bin\"
Public Const vc_cmdFfn$ = vc_binPth & "Code.Cmd"
Public Const vc_clijsPth$ = vc_Pth & "Resources\app\out\"
Public Const vc_clijsFfn$ = vc_clijsPth & "Cli.js"
Private Sub B_VcFt(): Vc "Abc": End Sub
Sub VcFt(Ft)
WrtACmd:
    Dim FfnACmd$: FfnACmd = PthTmpFdr("VcFt") & "A.cmd"
    Dim CxtACmd$: CxtACmd = FmtQQ("Code.Cmd ""?""", Ft)
    WrtStr CxtACmd, FfnACmd, OvrWrt:=True
ShellMax FfnACmd
End Sub
Sub NoteFt(Ft): ShellMax FmtQQ("notepad.exe ""?""", Ft): End Sub
Sub BrwFt(Ft, Optional UseVc As Boolean)
If UseVc Then
    VcFt Ft
Else
    NoteFt Ft
End If
End Sub

Function vc_CxtCodeCmd$(): vc_CxtCodeCmd = LinesFt(vc_cmdFfn):   End Function
Function vc_CxtCliJs$():     vc_CxtCliJs = LinesFt(vc_clijsFfn): End Function

Private Sub WW_NotUse()
'C:\Users\user\AppData\Local\Temp\JC\N20200820_213641.txt

'WrtAy MthnyPubPC, "C:\Users\user\AppData\Local\Temp\JC\N20200820_213641.txt"
'VcFt "C:\Users\user\AppData\Local\Temp\JC\N20200820_213641.txt"

'Vc "Ft"
'"C:\Users\user\AppData\Local\Programs\Microsoft VS Code\Code.exe" "C:\Users\user\AppData\Local\Temp\JC\N20200820_213641.txt"
'Cmd /c """C:\Users\user\AppData\Local\Programs\Microsoft VS Code\bin\Code1.cmd"" ""C:\Users\user\AppData\Local\Temp\JC\N20200820_213641.txt"""

'"C:\Users\user\AppData\Local\Programs\Microsoft VS Code\Code.exe" "C:\Users\user\AppData\Local\Programs\Microsoft VS Code\resources\app\out\cli.js"
'BrwFt "C:\Users\user\AppData\Local\Programs\Microsoft VS Code\bin\Code.cmd.backup"

'@echo off
'setlocal
'set VSCODE_DEV=
'Set ELECTRON_RUN_AS_NODE = 1
'"%~dp0..\Code.exe" "%~dp0..\resources\app\out\cli.js" %*
'endlocal

End Sub
Private Sub B_BrwHtml()
GoSub T1
Exit Sub
Dim Html$
T1:
    Html = "<html>lksdjflksdfjk<br>slkdfjlsdkjf<br></html>"
    GoTo Tst
Tst:
    BrwHtml Html
    Return
End Sub
Sub BrwHtml(Html$, Optional PfxFn$): BrwFtHtml FfnTmpCxt(".html", Html, "BrwHtml", PfxFn): End Sub
Sub BrwFtHtml(FtHtml$):              ShellFfnWin FtHtml:                                   End Sub
