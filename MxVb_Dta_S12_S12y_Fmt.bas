Attribute VB_Name = "MxVb_Dta_S12_S12y_Fmt"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_S12_S12y_Fmt."
Sub BrwS12y(A() As S12): BrwAy FmtS12y(A): End Sub

Private Sub B_FmtS12y()
Const ResPseg$ = "MxVbDtaS12Fmt\"
Dim S12y() As S12, H12$, Fmt As eTblFmt, Wdt%
GoSub T0
'GoSub T1
'GoSub Z1
'GoSub Z2
'GoSub T3
Exit Sub
T3:
    H12 = "AA BB"
    Ept = WTEpt(3)
    S12y = WTInp(3)
    H12 = "AA BB"
    Fmt = eTblFmtTb
    GoTo ZZ
    GoTo Tst
T0:
    S12y = S12yAdd(S12("S12y", "B"), S12("AA", "B"))
    H12 = "AA BB"
    Fmt = eTblFmtTb
    GoTo ZZ
Z2:
    S12y = WTSamp2
    H12 = "AA BB"
    Fmt = eTblFmtTb
    GoTo ZZ
Z1:
    H12 = "AA BB"
    S12y = WTSamp1
    Fmt = eTblFmtSS
    GoTo ZZ
ZZ:
    Brw FmtS12y(S12y, H12, Wdt)
    Return
Tst:
    Act = FmtS12y(S12y, H12)
    C
    Return
End Sub
Function PmsS12$(S As S12): PmsS12 = "-S1 " & QuoTm(S.S1) & "-S2 " & QuoTm(S.S2)
End Function
Function FmtS12y(A() As S12, Optional H12$ = "S1 S2", Optional Wdt% = 120) As String()
Dim D As Drs: D = Drs(FnyFF(H12), WDyS12y(A))
FmtS12y = FmtDrs(D, eRdsNo, eTblFmtTb, eZerShw, Wdt)
End Function
Private Function WDyS12y(A() As S12) As Variant()
Dim J&: For J = 0 To UbS12(A)
    With A(J)
        PushI WDyS12y, Array(.S1, .S2)
    End With
Next
End Function

Private Sub WTEdtInp(Cas%):                            EdtRES X_InpSegFn(Cas):            End Sub
Private Sub WTEdtEpt(Cas%):                            EdtRES X_EptSegFn(Cas):            End Sub
Private Function WTInp(Cas%) As S12():         WTInp = resS12y(WTInpSegFn(Cas)):          End Function
Private Function WTEpt(Cas%) As String():      WTEpt = Resy(WTEptSegFn(Cas)):             End Function
Private Function WTInpSegFn$(Cas%):       WTInpSegFn = WTCasPseg(Cas) & "Inp.txt":        End Function
Private Function WTEptSegFn$(Cas%):       WTEptSegFn = WTCasPseg(Cas) & "Ept.txt":        End Function
Private Function WTCasPseg$(Cas%):         WTCasPseg = FmtQQ("MxVbDtaS12Fmt\Cas?\", Cas): End Function
Private Function WTSamp1() As S12()
PushS12 WTSamp1, S12("sldjflsdkjf", "lksdjf")
PushS12 WTSamp1, S12("sldjflsdkjf", "lksdjf")
PushS12 WTSamp1, S12("sldjf", "lksdjf")
PushS12 WTSamp1, S12("sldjdkjf", "lksdjf")
End Function
Private Function WTSamp2() As S12()
PushS12 WTSamp2, S12(RplVBar("sldj|flsdk|jf"), RplVBar("lk|sdjf"))
PushS12 WTSamp2, S12("sldjflsdkjf", RplVBar("lksdj|f"))
PushS12 WTSamp2, S12("sldjf", "lksdjf")
PushS12 WTSamp2, S12("sldjdkjf", "lksdjf")
End Function

Private Function X_InpSegFn$(Cas%): X_InpSegFn = W5Cas(Cas) & "Inp.txt":            End Function
Private Function X_EptSegFn$(Cas%): X_EptSegFn = W5Cas(Cas) & "Ept.txt":            End Function
Private Function W5Cas$(Cas%):           W5Cas = FmtQQ("MxVbDtaS12Fmt\Cas?\", Cas): End Function
