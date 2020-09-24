Attribute VB_Name = "MxIde_Mth_Drs_TMthCnt"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_Drs_TMthCnt."
Const CntgMthPP$ = "NPSub NPFun NPPrp NPrvSub NPrvFun NPrvPrp NFrdSub NFrdFun NFrdPrp"
Public Const MthCntFf$ = "Pjn Mdn NLn NMth NPSub NPFun NPPrp NPrvSub NPrvFun NPrvPrp NFrdSub NFrdFun NFrdPrp"
Type MthCnt
    Pjn As String
    Mdn As String
    NLn As String
    NMth As Integer
    NPSub As Integer
    NPFun As Integer
    NPPrp As Integer
    NPrvSub As Integer
    NPrvFun As Integer
    NPrvPrp As Integer
    NFrdSub As Integer
    NFrdFun As Integer
    NFrdPrp As Integer
End Type

Function DrsTMthCntPC(Optional PatnssAndMd$, Optional SrtCol$ = "Mdn") As Drs: DrsTMthCntPC = DrsTMthCntP(CPj, PatnssAndMd, SrtCol): End Function
Function DrsTMthCntP(P As VBProject, PatnssAndMd$, SrtCol$) As Drs
Dim Dy()
    Dim R As RegExp: Set R = Rx(PatnssAndMd)
    Dim C As VBComponent: For Each C In P.VBComponents
        If R.Test(C.Name) Then
            PushI Dy, WDr(C.CodeModule)
        End If
    Next
Dim D As Drs: D = DrsFf(MthCntFf, Dy)
DrsTMthCntP = DrsSrt(D, SrtCol)
End Function
Function DrsTMthCntM(M As CodeModule) As Drs: DrsTMthCntM = DrsFf(MthCntFf, WDr(M)): End Function
Private Function WDr(M As CodeModule) As Variant()
With WMthCnt(M)
    WDr = Array(.Pjn, .Mdn, .NLn, .NMth, .NPSub, .NPFun, .NPPrp, .NPrvSub, .NPrvFun, .NPrvPrp, .NFrdSub, .NFrdFun, .NFrdPrp)
End With
End Function
Private Function WMthCnt(M As CodeModule) As MthCnt: WMthCnt = WMthCntSrc(SrcM(M), Mdn(M), PjnM(M)): End Function
Private Function WMthCntSrc(Src$(), Mdn$, Pjn$) As MthCnt
Const CSub$ = CMod & "WMthCntSrc"
With WMthCntSrc
Dim Mthlny$(): Mthlny = MthlnySrc(Src)
Dim L: For Each L In Itr(Mthlny)
    With TMthL(L)
        Dim Prv As Boolean: Prv = False
        Dim Pub As Boolean: Pub = False
        Dim Frd As Boolean: Frd = False
        Dim Fun As Boolean: Fun = False
        Dim Sbr As Boolean: Sbr = False
        Dim Prp As Boolean: Prp = False
        Select Case .ShtMdy
        Case "Prv": Prv = True
        Case "Pub", "": Pub = True
        Case "Frd": Frd = True
        Case Else: Thw CSub, "Out of valid value: Prv Pub Frd", "ShtMdy", .ShtMdy
        End Select
        Select Case Shtmthkd(.ShtTy)
        Case "Fun": Fun = True
        Case "Sub": Sbr = True
        Case "Prp": Prp = True
        Case Else: Thw CSub, "Out of valid value: Sub Fun Prp", "ShtMdy", .ShtMdy
        End Select
    End With
            
    Select Case True
        Case Pub And Sbr: .NPSub = .NPSub + 1
        Case Pub And Fun: .NPFun = .NPFun + 1
        Case Pub And Prp: .NPPrp = .NPPrp + 1
        Case Prv And Sbr: .NPrvSub = .NPrvSub + 1
        Case Prv And Fun: .NPrvFun = .NPrvFun + 1
        Case Prv And Prp: .NPrvPrp = .NPrvPrp + 1
        Case Frd And Sbr: .NFrdSub = .NFrdSub + 1
        Case Frd And Fun: .NFrdFun = .NFrdFun + 1
        Case Frd And Prp: .NFrdPrp = .NFrdPrp + 1
        Case Else: Thw CSub, "Invalid TMth", "Mthln", L
    End Select
    .NMth = .NMth + 1
    If .NPSub + .NPFun + .NPPrp + .NPrvSub + .NPrvFun + .NPrvPrp + .NFrdSub + .NFrdFun + .NFrdPrp <> .NMth Then Stop
Next
.NLn = Si(Src)
.Mdn = Mdn
.Pjn = Pjn
End With
End Function

Function NMthM%(M As CodeModule): NMthM = NMthSrc(SrcM(M)): End Function
Function NSrcLinPj&(P As VBProject)
Dim O&, C As VBComponent
For Each C In P.VBComponents
    O = O + C.CodeModule.CountOfLines
Next
NSrcLinPj = O
End Function
Function NMthPub%(Src$()):            NMthPub = Si(MthnyPub(Src)): End Function
Function NMthPubM%(M As CodeModule): NMthPubM = NMthSrc(SrcM(M)):  End Function
Function NMthPubV%(V As VBE)
Dim P As VBProject: For Each P In V.VBProjects
    NMthPubV = NMthPubV + NMthPubP(P)
Next
End Function
Function NMthPubVC%(): NMthPubVC = NMthPubV(CVbe): End Function
Function NMthPubP%(P As VBProject)
Dim O%, C As VBComponent
For Each C In P.VBComponents
    O = O + NMthPubM(C.CodeModule)
Next
NMthPubP = O
End Function
Function NMth%(A As MthCnt)
With A
NMth = .NPSub + .NPFun + .NPPrp + .NPrvSub + .NPrvFun + .NPrvPrp + .NFrdSub + .NFrdFun + .NFrdPrp
End With
End Function
Function NMthSrc%(Src$()): NMthSrc = Si(Mthixy(Src)): End Function
Function NMthPC%():         NMthPC = NMthP(CPj):      End Function
Function NMthMC%():         NMthMC = NMthM(CMd):      End Function
Function NMthP%(Pj As VBProject)
Dim O%, C As VBComponent
For Each C In Pj.VBComponents
    O = O + NMthSrc(SrcM(C.CodeModule))
Next
NMthP = O
End Function

Function FmtCntgMth(A As MthCnt, Optional Hdr As eHdr)
Dim Pfx$: If Hdr = eHdrNo Then Pfx = "Pub* | Prv* | Frd* : *{Sub Fun Frd} "
With A
Dim N%: N = NMth(A)
FmtCntgMth = JnApSep(" | ", Pfx, .Mdn, N) & " | " & JnSpcAp(.NPSub, .NPFun, .NPPrp, .NPrvSub, .NPrvFun, .NPrvPrp, .NFrdSub, .NFrdFun, .NFrdPrp)
End With
End Function
