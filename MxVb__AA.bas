Attribute VB_Name = "MxVb__AA"
Option Compare Text
Option Explicit

Function P12SepTab(S) As P12: P12SepTab = P12Sep(S, vbTab): End Function
Function P12Sep(S, Sep$, Optional C As eCas) As P12
Dim P1%: P1 = PosSsub(S, Sep, C)
If P1 = 0 Then Exit Function
Dim P2%: P2 = PosSsub(S, Sep, C, P1 + Len(Sep))
If P2 = 0 Then Thw CSub, "No 2nd @Sep", "@S @Sep @C", S, Sep, EnmsCas(C)
P12Sep = P12(P1, P2)
End Function


Function FfnNxtAva$(Ffn) 'MxVb_Fs_Ffn_FfnNxt
Const CSub$ = CMod & "FfnNxtAva"
Dim J%, O$
O = Ffn
While HasFfn(O)
    If J = 9999 Then Thw CSub, "Too much next file in the path of given-ffn", "Given-Ffn", Ffn
    J = J + 1
    O = FfnNxt(O)
Wend
FfnNxtAva = O
End Function


Function FnNxt$(Fn$, Fnay$(), Optional MaxN% = 999) 'MxVb_Fs_Ffn_FfnNxt
If Not HasEle(Fnay, Fn) Then FnNxt = Fn: Exit Function
FnNxt = EleMax(AwLik(Fnay, Fn & "(????)"))
End Function


Function IsFfnNxt(Ffn) As Boolean 'MxVb_Fs_Ffn_FfnNxt
'Select Case True
'Case WNbrNxt(Ffn) > 0, Right(Fnn(Ffn), 5) = "(000)": IsFfnNxt = True
'End Select
End Function
