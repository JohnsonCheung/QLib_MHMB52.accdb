Attribute VB_Name = "MxIde_Ctl_NoPm"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Ctl_NoPm."

Function IBarny() As String():     IBarny = Itn(IBarsC):                       End Function
Function NWin&():                    NWin = CVbe.Windows.Count:                End Function
Function NWinVis&():              NWinVis = NItpTrue(CVbe.Windows, "Visible"): End Function
Function CapyIWin() As String(): CapyIWin = SyItp(CVbe.Windows, "Caption"):    End Function
Function CapyIWinVis() As String()
Dim W: W = IWinyVis
CapyIWinVis = SyOyP(W, "Caption")
End Function

Function IWinyVis() As VbIde.Window()
Dim W As VbIde.Window: For Each W In CVbe.Windows
    If W.Visible Then PushObj IWinyVis, W
Next
End Function
Function WinnyVis() As String(): WinnyVis = Itn(IWinyVis): End Function

Function sampspBtn() As String()
Erase XX
X "Bars"
X " AA A1 A2 A3"
X " BB B1 B2 B3"
X "Btns"
X " A1"
sampspBtn = XX
End Function
