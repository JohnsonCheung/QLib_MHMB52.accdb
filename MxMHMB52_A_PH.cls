VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_PH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_PH."

Property Get Ss5$():                              Ss5 = "L1 L2 L3 L4 LSku":       End Property
Property Get Ss7$():                              Ss7 = "LStm LBus " & Ss5:       End Property
Function Tmpy7Qtp(QtpTb$) As String(): Tmpy7Qtp = AmRplQ(Itmy7, QtpTb):           End Function
Property Get Tmpy7ScOH() As String():             Tmpy7ScOH = Tmpy7Qtp("$ScOH?"): End Property
Property Get Tmpy7Days() As String():             Tmpy7Days = Tmpy7Qtp("$Days?"): End Property
Property Get Itmy7() As String():                 Itmy7 = SplitSpc(Ss7):          End Property
Property Get Tmpy7() As String():                 Tmpy7 = NyQtp2("$?", Ss7):      End Property
Private Sub DrpTmpy7(QtpTb$):                     DrpTnyC Tmpy7Qtp(QtpTb):        End Sub
Sub DrpTmpy7Days():                               DrpTmpy7 "$Days?":              End Sub
Sub DrpTmpy7ScOH():                               DrpTmpy7 "$ScOH?":              End Sub
