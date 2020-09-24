Attribute VB_Name = "MxVb_Run"
'Declare Function GetProcessId& Lib "Kernel32.dll" (ProcessHandle&)
'Const Ps1Str$ = "function Get-ExcelProcessId { try { (Get-Process -Name Excel).Id } finally { @() } }" & vbCrLf & _
'"Stop-Process -Id (Get-ExcelProcessId)"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Run."
Enum eWaitRslt: eTimUp: eCnl: End Enum
Type WaitOpt
    TimOutSec As Integer
    ChkSec As Integer
    KeepFcmd As Boolean
End Type
Declare Function GetCurrentProcessId& Lib "Kernel32.dll" ()

Function WaitOpt(TimOutSec%, ChkSec%, KeepFcmd As Boolean) As WaitOpt
With WaitOpt
.TimOutSec = TimOutSec
.ChkSec = ChkSec
.KeepFcmd = KeepFcmd
End With
End Function

Function DftWaitOpt() As WaitOpt: DftWaitOpt = WaitOpt(30, 5, False): End Function

Function RunFps1&(Fps1$, Optional PmStr$, Optional Sty As VbAppWinStyle = vbMaximizedFocus): RunFps1 = RunFcmd("PowerShell", QuoDbl(Fps1) & " " & PmStr, Sty): End Function
Function RunFcmd&(Fcmd$, Optional PmStr$, Optional Sty As VbAppWinStyle = vbMaximizedFocus): RunFcmd = Shell(QuoDbl(Fcmd) & " " & PmStr, Sty):                 End Function

Sub ChkWaitFfn(Ffn$, Optional Fun$, Optional ChkSec% = 1, Optional TimOutSec% = 5, Optional Sty As VbAppWinStyle = VbAppWinStyle.vbMaximizedFocus, Optional IsKillFfn As Boolean)
If WaitFfnExi(Ffn, ChkSec, TimOutSec, Sty, IsKillFfn) Then
    Wait  ' wait a second after ffn is found existed
    Exit Sub
End If
Thw Fun, "Timeout to wait for a file to be created", "[File Path] [File Name] ChkSec TImOutSec", Pth(Ffn), Fn(Ffn), ChkSec, TimOutSec
End Sub
Function WaitFfnExi(Ffn$, Optional TimOutSec% = 5, Optional ChkSec% = 1, Optional Sty As VbAppWinStyle = VbAppWinStyle.vbMaximizedFocus, Optional IsKillFfn As Boolean) As Boolean _
'Return True, if Ffn is exist/created within TimOutSec otherwise false
Dim J%: For J = 1 To TimOutSec \ ChkSec
    If HasFfn(Ffn) Then
        If IsKillFfn Then Kill Ffn
        Exit Function
    End If
    If Not Wait Then Exit Function
Next
WaitFfnExi = True
End Function

Private Sub B_RunFcmd()
RunFcmd "Cmd"
MsgBox "AA"
End Sub

Function Wait(Optional Sec% = 1) As eWaitRslt
Dim Till As Date: Till = AftSec(Sec)
Wait = IIf(Xls.Wait(Till), eTimUp, eCnl)
End Function

Function AftSec(Sec%) As Date 'Return the Date after Sec from Now
AftSec = DateAdd("S", Sec, Now)
End Function
Function Pipe(Pm, Mthnn$)
Dim O: Asg Pm, O
Dim I
For Each I In SySs(Mthnn)
   Asg Run(I, O), O
Next
Asg O, Pipe
End Function

Function RunAvIgnEr(Mthn, Av())
Const CSub$ = CMod & "RunAvIgnEr"
If Si(Av) > 9 Then Thw CSub, "Si(Av) should be 0-9", "Si(Av)", Si(Av)
On Error Resume Next
RunAv Mthn, Av
End Function

Function RunAv(Mthn, Av())
Const CSub$ = CMod & "RunAv"
Dim O
Select Case Si(Av)
Case 0: O = Run(Mthn)
Case 1: O = Run(Mthn, Av(0))
Case 2: O = Run(Mthn, Av(0), Av(1))
Case 3: O = Run(Mthn, Av(0), Av(1), Av(2))
Case 4: O = Run(Mthn, Av(0), Av(1), Av(2), Av(3))
Case 5: O = Run(Mthn, Av(0), Av(1), Av(2), Av(3), Av(4))
Case 6: O = Run(Mthn, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5))
Case 7: O = Run(Mthn, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5), Av(6))
Case 8: O = Run(Mthn, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5), Av(6), Av(7))
Case 9: O = Run(Mthn, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5), Av(6), Av(7), Av(8))
Case Else: Thw CSub, "UB-Av should be <= 8", "UB-Si Mthn", UB(Av), Mthn
End Select
RunAv = O
End Function
