Attribute VB_Name = "MxIde_MthSts"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_MthSts."
Type TMthSts
    Mdn As String
    NSrcln As Long
    NPub As Integer
    NPrv As Integer
    NFrd As Integer
    NPrvWHdr As Integer ' W_{N}
    NPrvWDet As Integer
    NPrvXHdr As Integer
    NPrvXDet As Integer
    NPrvZ As Integer
End Type

Function LnTMthSts$(S As TMthSts)
With S
LnTMthSts = FmtQQ("[NSrcLn NPub NPrv NFrd](? ? ? ?)", .NSrcln, .NPub, .NPrv, .NFrd)
End With
End Function

Sub DmpTMthStsMC():               DmpTMthStsM CMd:   End Sub
Sub DmpTMthStsM(M As CodeModule): W2Dmp TMthStsM(M): End Sub
Private Sub W2Dmp(S As TMthSts):  D LnTMthSts(S):    End Sub

Function TMthStsMC() As TMthSts:                TMthStsMC = TMthStsM(CMd):                          End Function
Function TMthStsM(M As CodeModule) As TMthSts:   TMthStsM = TMthStsSrc(Mdn(M), SrcM(M)):            End Function
Function TMthStsSrc(Mdn$, Src$()) As TMthSts:  TMthStsSrc = WTMthSts(Mdn, Si(Src), MthlnySrc(Src)): End Function
Private Function WTMthSts(Mdn$, NSrcln&, MthlnySrc$()) As TMthSts
With WTMthSts
    .Mdn = Mdn
    .NSrcln = NSrcln
    Dim L: For Each L In Itr(MthlnySrc)
        Select Case Mdy(L)
        Case "", "Public": .NPub = .NPub + 1
        Case "Private":    .NPrv = .NPrv + 1
        Case "Friend":     .NFrd = .NFrd + 1
        Case Else: Stop
        End Select
    Next
End With
End Function
