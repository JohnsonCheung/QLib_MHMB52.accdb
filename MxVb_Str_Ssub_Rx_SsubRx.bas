Attribute VB_Name = "MxVb_Str_Ssub_Rx_SsubRx"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Ssub_Rx_SsubRx."
Function SsubyMchColl(M As MatchCollection) As String(): SsubyMchColl = SyItv(M):                       End Function
Function IyMchColl(M As MatchCollection) As Integer():      IyMchColl = IntyItp(M, "FirstIndex"):       End Function
Function SsubPatn$(S, Patn$, Optional IthSubMch%):           SsubPatn = SsubRx(S, Rx(Patn), IthSubMch): End Function
Function SsubRx$(S, Rx As RegExp, Optional IthSubMch%) ' ret @IthSubMch-mch of fst-mch.  Thw if the @IthSubMch > N-Sub-Mch
'fst-mch: it is fst mch of the Mch-Coll.  If no fst-mch, return ""
'if fst-mch is found, return the SsubPatn of the @IthSubMch, but if @IthSubMch is greater the N-Sub-Mch of the fst-mch, thw error

Dim M As MatchCollection: Set M = Rx.Execute(S)
If M.Count = 0 Then Exit Function ' No match
If IthSubMch = 0 Then
     Dim M0 As Match: Set M0 = M(0): ' Fst Matches
    SsubRx = M0.Value
    Exit Function
End If
Dim SMch As SubMatches: Set SMch = WSMchRx(S, Rx)
SsubRx = WSsubSMch(SMch, IthSubMch)
End Function
Private Function WSMchRx(S, Rx As RegExp) As SubMatches
Dim M As Match: Set M = ItvFst(Rx.Execute(S)): If IsNothing(M) Then Exit Function
Set WSMchRx = M.SubMatches
End Function


Private Sub B_SsubyRx(): BrwAy SsubyRx("#AA# #BB# #A A#", Rx("/#(\w[\w:-]*)#/G")): End Sub
Function SsubyPatn(S, Patn$, Optional IthSubMch%) As String()  ' Sy of Ssub of each ele of @Sy by @Patn
SsubyPatn = SsubyRx(S, Rx(Patn), IthSubMch)
End Function
Function SsubyRx(S, R As RegExp, Optional IthSubMch%) As String()
If IthSubMch = 0 Then
    SsubyRx = SyItv(Mchcoll(S, R)) ' Sy of (All Ssub of each ele of @Sy by @Rx)
Else
    Dim M As Match: For Each M In Mchcoll(S, R)
        PushI SsubyRx, M.SubMatches(IthSubMch - 1)
    Next
End If

End Function
Private Function WSsubSMch$(S As SubMatches, IthSMch%) 'return [Str]-[Mch]Matched from [Sub]Matches-@S of @IthSMch
Const CSub$ = CMod & "WSsubSMch"
With S
    If .Count < IthSMch Then ThwPm CSub, "@SubMatches-count is less thatn @IthSMch", "@S-SMchCnt @IthSMch", S.Count, IthSMch
    If .Count = 0 Then ThwPm CSub, "Given @SMch should have count>0"
    WSsubSMch = .Item(IthSMch - 1)
End With
End Function
