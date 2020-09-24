Attribute VB_Name = "MxVb_Fs_DInp"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_DInp."

Function sampspLimInp() As String()
Erase XX
X "MB52 C:\Users\user\Desktop\Mhd\SAPAccessReports\StockShipCost\Sample\MB52 2018-07-30.xls"
X "UOM  C:\Users\user\Desktop\Mhd\SAPAccessReports\StockShipCost\Sample\sales text.xlsx"
X "ZHT1 C:\Users\user\Desktop\Mhd\SAPAccessReports\StockShipCost\Sample\ZHT1.XLSX"
sampspLimInp = XX
End Function

Function MsgDrs(Msg$, A As Drs) As String()
Erase XX
XLn Msg
XDrs A
XLn
MsgDrs = XX
End Function

Private Sub B_ErSplimInp(): Brw W1ErSpLnLimInp(sampspLimInp): End Sub
Private Function W1ErSpLnLimInp(SpLnLimInp$()) As String()
Dim E1$(), E2$(), E3$(), E4$()
Dim Lnoy%(), Inpny$(), Ffny$(), FilKdy$()
AsgT3ry SpLnLimInp, Lnoy, Inpny, FilKdy, Ffny
E1 = WErDup(Lnoy, Inpny, "Inpn")
E2 = WErDup(Lnoy, Ffny, "Ffn")
E3 = W1EryFfnMis(Lnoy, Ffny)
E4 = W1ErFilKd(Lnoy, FilKdy)
W1ErSpLnLimInp = Sy(E1, E2, E3)
End Function
Private Function W1ErFilKd(Lnoy%(), FilKdy$()) As String()

End Function
Private Function WErDup(Lnoy%(), Sy$(), Dtan$) As String()

End Function
Private Function W1EryFfnMis(Lnoy%(), Ffny$()) As String()
Dim I%: Stop 'I = IxEle(WiFfn.Fny, "Ffn")
Dim Dr, Dy(): 'For Each Dr In Itr(WiFfn.Dy)
    If NoFfn(Dr(I)) Then PushI Dy, Dr
'Next
Dim B As Drs: Stop 'B = Drs(WiFfn.Fny, Dy)
Stop 'W1EryFfnMis = MsgDrs("File not exist", B)
End Function
Private Function W1ErFilKdy(Lnoy%(), FilKdy$()) As String()

End Function
