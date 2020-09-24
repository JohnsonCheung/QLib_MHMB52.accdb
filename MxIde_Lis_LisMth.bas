Attribute VB_Name = "MxIde_Lis_LisMth"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Lis_LisMth."
Enum eWhMdy: eWhMdyPub: eWhMdyPrv: eWhMdyAll: End Enum: Public Const EnmmMdy$ = "eWhMdy? Pub Prv All"
Public Const EnmqssWhMdy$ = "eWhMdy? Pub Prv All"
Private Sub B_DmpMth(): DmpMth , "CCml": End Sub
Function HitMdy(ByVal Mdy, W As eWhMdy) As Boolean
If Mdy = "" Then Mdy = "Pub"
Select Case True
Case W = eWhMdyPrv: HitMdy = Mdy = "Prv"
Case W = eWhMdyPub: HitMdy = Mdy = "Pub"
Case W = eWhMdyAll: HitMdy = True
Case Else: ThwEnm CSub, W, EnmqssWhMdy
End Select
End Function
Sub DmpMth(Optional PatnssAndMd$, Optional SsMthPatn$, Optional IsSrtByMthn As Boolean, Optional M As eWhMdy): DmpAy WFmtLisMth(PatnssAndMd, SsMthPatn, IsSrtByMthn, M): End Sub
Sub BrwMth(Optional PatnssAndMd$, Optional SsMthPatn$, Optional IsSrtByMthn As Boolean, Optional M As eWhMdy): BrwAy WFmtLisMth(PatnssAndMd, SsMthPatn, IsSrtByMthn, M): End Sub
Sub VcMth(Optional PatnssAndMd$, Optional SsMthPatn$, Optional IsSrtByMthn As Boolean, Optional M As eWhMdy):  VcAy WFmtLisMth(PatnssAndMd, SsMthPatn, IsSrtByMthn, M):  End Sub
Sub LisMth(Optional PatnssAndMd$, Optional SsMthPatn$, Optional IsSrtByMthn As Boolean, Optional M As eWhMdy): DmpAy WFmtLisMth(PatnssAndMd, SsMthPatn, IsSrtByMthn, M): End Sub

Private Function WFmtLisMth(PatnssAndMd$, SsMthPatn$, IsSrtByMthn As Boolean, M As eWhMdy) As String()
Dim R() As RegExp: R = Rxay(SsMthPatn)
Dim O$()
    Dim Mdny$(): Mdny = SySrtQ(MdnyP(CPj, PatnssAndMd))
    Dim N: For Each N In Itr(Mdny)
        DoEvents
        PushIAy O, WMthnyCmp(Cmp(N), R, M)
    Next
Dim HypKeyCii$: If IsSrtByMthn Then HypKeyCii = "1"
WFmtLisMth = FmtT1ry(O, , HypKeyCii, NoIx:=True)
End Function
Private Function WMthnyCmp(C As VBComponent, Rxay() As RegExp, M As eWhMdy) As String()
Dim S$(): S = SrcCmp(C): If Si(S) = 0 Then Exit Function
Dim N$()
    N = MthnWhMdy(S, M)
    N = SySrtQ(AwRxAyAnd(N, Rxay))
    If False Then
        Dim Mthn: For Each Mthn In N
            If HasSsub(Mthn, "SlmCpr") Then Stop
        Next
    End If
WMthnyCmp = AmAddPfx(N, C.Name & " ")
End Function
