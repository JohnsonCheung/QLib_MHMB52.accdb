Attribute VB_Name = "MxVb_Run_StopXls"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Run_StopXls."

Sub StopXls()
Dim F$: F = PthTmp & "StopXls.Ps1"
EnsFt F, WCdl
Shell F, vbMaximizedFocus
End Sub

Private Function WCdl$()
WCdl = LinesApLn( _
    "PowerShell -Command ""try{Stop-Process -Id{try{(Get-Process -Name Excel).Id}finally{}.invoke()}""", _
    "Pause...")
End Function
