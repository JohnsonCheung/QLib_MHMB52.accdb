Attribute VB_Name = "MxIde_Mthn_Cml_Mthnvanv"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthn_Cml_MthMthvnav."
Sub VcMthiyVnavMthln(): Vc FmtT4ry(MthiyVnavMthlnPubPC, NoIx:=True), "pub verb noun adje var Mthny": End Sub
Sub VcMthiyNavvMthln(): Vc FmtT4ry(MthiyNavvMthlnPubPC, NoIx:=True), "pub noun adje verb var Mthny": End Sub
Function MthiyNavvMthlnPubPC() As String()
Dim Mthln: For Each Mthln In MthlnyPubPC
    Dim Mi2TyNm$: Mi2TyNm = Mi2TyNmTMth(TMthL(Mthln))
    PushI MthiyNavvMthlnPubPC, Mi4NavvzVnavv(Mi4VnavzM2(Mi2TyNm)) & " " & Mthln
Next
End Function
Function MthiyVnavMthlnPubPC() As String()
Dim Mthln: For Each Mthln In MthlnyPubPC
    Dim Mi2TyNm$: Mi2TyNm = Mi2TyNmTMth(TMthL(Mthln))
    PushI MthiyVnavMthlnPubPC, Mi4VnavzM2(Mi2TyNm) & " " & Mthln
Next
End Function

Private Function Mi4NavvzVnavv$(Mthvnav$)
Dim A$(): A = SplitSpc(Mthvnav)
Dim U%: U = UB(A): If U <> 3 Then Stop
Mi4NavvzVnavv = RTrim(A(1) & " " & A(2) & " " & A(0) & " " & A(3))
End Function
Function Mi4NavvFun$(Funn$):  Mi4NavvFun = Mi4NavvzVnavv(Mi4VnavzVerbNav(".Fnd", Funn)): End Function
Function Mi4NavvM2$(Mi2TyNm):  Mi4NavvM2 = Mi4NavvzVnavv(Mi4VnavzM2(Mi2TyNm)):           End Function
Function Mi4VnavzM2$(Mi2TyNm)
Dim ShtTy$, Mthn$
AsgT1r Mi2TyNm, ShtTy, Mthn
Select Case ShtTy
Case "Fun": Mi4VnavzM2 = Mi4VnavzVerbNav(".Fnd", Mthn)
Case "Get": Mi4VnavzM2 = Mi4VnavzVerbNav(".Get", Mthn)
Case "Sub": Mi4VnavzM2 = Mi4VnavzSub(Mthn)
Case "Set": Mi4VnavzM2 = Mi4VnavzVerbNav("Set", Mthn)
Case "Let": Mi4VnavzM2 = Mi4VnavzVerbNav("Let", Mthn)
Case Else: ThwPm CSub, "Fst term of @Mi2TyNm must be [Fun Get Sub Set Let]", "@Mi2TyNm", Mi2TyNm
End Select
End Function

Private Function Mi4VnavzVerbNav$(Verb$, Nav$)
Dim Adje$, Noun$, Var$
AsgFunn Nav, Noun, Adje, Var
Mi4VnavzVerbNav = Mi4VnavzVerb(Verb, Noun, Adje, Var)
End Function

Private Sub AsgSubn__Tst()
GoSub T1
Exit Sub
Dim Mthvnav$, oVerb$, oAdje$, oNoun$, oVar$
T1:
    Mthvnav = "RaiseQQ"
    GoTo Tst
Tst:
    AsgSubn Mthvnav, oVerb, oAdje, oNoun, oVar
    Return
End Sub
Private Sub AsgSubn(Subn$, oVerb$, oNoun$, oAdje$, oVar$)
Dim M$
Dim S$: S = Subn
oVar = ShfMiVar(S)
oVerb = ShfCCmlFst(S): If oVerb = "" Then M = "Verb cannot be blank": GoTo Thw
oNoun = StrDotIf(ShfCCmlFst(S)) ' Put . if blank
oAdje = StrDotIf(S)             ' Put . if blank
Exit Sub
Thw: Thw CSub, M, "[Mthn which should be able to break in Verb-Noun-Adje-Var]", Subn
End Sub
Private Sub ShfMiVar__Tst()
GoSub T1
GoSub T2
GoSub T3
Exit Sub
Dim OMthn$, EptOMthn$
T1:
    OMthn = "AAAzAA"
    Ept = "zAA"
    EptOMthn = "AAA"
    GoTo Tst
T2:
    OMthn = "AAAAA"
    Ept = "."
    EptOMthn = "AAAAA"
    GoTo Tst
T3:
    OMthn = "DrsFxq"
    Ept = "."
    EptOMthn = "DrsFxq"
    GoTo Tst
Tst:
    Act = ShfMiVar(OMthn)
    Ass EptOMthn = OMthn
    C
    Return
End Sub
Private Function ShfMiVar$(OMthn$)
'#Mthnvar:Mthn-Variant# It is optional.  It is Sfx of mthn in 2 cases. Case1: All UCas. Case2: after letter z
Dim P%: P = PosMiz(OMthn)

If P > 0 Then
    ShfMiVar = Mid(OMthn, P)
    OMthn = Left(OMthn, P - 1)
    Exit Function
End If
For P = Len(OMthn) To 1 Step -1
    If Not IsUCas(Mid(OMthn, P, 1)) Then
        ShfMiVar = StrDotIf(Mid(OMthn, P + 1))
        OMthn = Left(OMthn, P)
        Exit Function
    End If
Next
ShfMiVar = "."
End Function
Private Function PosMiz%(Mthn$)
Dim PosBeg%: PosBeg = 1
L:
    Dim P%: P = PosSsub(Mthn, "z", eCasSen, PosBeg): If P = 0 Then Exit Function
    If P = Len(Mthn) Then Exit Function
    If IsUCas(Mid(Mthn, P + 1, 1)) Then PosMiz = P: Exit Function
    PosBeg = P + 1
    GoTo L
End Function

Private Sub AsgFunn(Funn$, oNoun$, oAdje$, oVar$)
'#Mthnav:Mth-Noun-Adje-Var# a mthn which can be broken into Noun-Adje-Avar
' The Mthn should be able to be broken into Noun-[Adje]-[Var].  in 2 cases: ChrFst is UCas or LCas
Dim Msg$
Dim S$: S = Funn
oVar = ShfMiVar(S)

If IsLCas(ChrFst(S)) Then
    Dim P%: P = PosChrUCasFst(S)
    If P = 0 Then Msg = "Fun-Mthn cannot have no noun.  That after remove Mthnvar, there should be CCml": GoTo Thw
    oAdje = Left(S, P - 1)
    oNoun = Mid(S, P)
Else
    oNoun = ShfCCmlFst(S)
    oAdje = S
End If
If oAdje = "" Then oAdje = "."
If oNoun = "" Then Msg = "Noun cannot be blank": GoTo Thw
Exit Sub
Thw: Thw CSub, Msg, "Mthn-should-be-[Noun-Adje-Var]", Funn
End Sub
Function Mi4VnavzSub$(Subn$)
Dim Verb$, Adje$, Noun$, Var$
AsgSubn Subn, Verb, Noun, Adje, Var
Mi4VnavzSub = Mi4VnavzVerb(Verb, Noun, Adje, Var)
End Function
Function Mi4VnavzLet$(Letn$)
Dim Adje$, Noun$, Var$
AsgFunn Letn, Noun, Adje, Var
Mi4VnavzLet = Mi4VnavzVerb("Let", Noun, Adje, Var)
End Function
Function Mi4VnavzSet$(Setn$)
Dim Adje$, Noun$, Var$
AsgFunn Setn, Noun, Adje, Var
Mi4VnavzSet = Mi4VnavzVerb("Set", Noun, Adje, Var)
End Function
Private Function Mi4VnavzVerb$(Verb$, Noun$, Adje$, Var$)
Mi4VnavzVerb = Verb & " " & Noun & " " & Adje & " " & Var
Select Case True
Case Verb = "", Adje = "", Noun = "", Var = "": Stop
End Select
End Function
Function Mi4NavvzMi2$(Mi2TyNm): Mi4NavvzMi2 = Mi4NavvzVnavv(Mi4VnavzM2(Mi2TyNm)): End Function
Function MiNounzMi2$(Mi2TyNm):   MiNounzMi2 = Tm1(Mi4NavvzMi2(Mi2TyNm)):          End Function
Function MiAdjelzFunn$(Funn)
If Not IsLCas(ChrFst(Funn)) Then Exit Function
Dim P%: P = PosChrUCasFst(Funn): If P = 0 Then Thw CSub, "Funn are all LCas which is not allowed", "Funn", Funn
MiAdjelzFunn = Left(Funn, P - 1)
End Function
