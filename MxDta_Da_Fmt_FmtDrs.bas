Attribute VB_Name = "MxDta_Da_Fmt_FmtDrs"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Fmt_FmtDrs."
Enum eRds: eRdsNo: eRdsYes: End Enum  'Deriving(Enm)
Enum eTblFmt: eTblFmtTb: eTblFmtSS: End Enum  'Deriving(Enm)
Enum eZer: eZerHid: eZerShw: End Enum 'Deriving(Enm)
Public Const EnmqssZer$ = "eZer? Hid Shw"
Public Const EnmqssTblFmt$ = "eTblFm? TB SS"
Function EnmsTblFmt$(E As eTblFmt): EnmsTblFmt = EnmsyTblFmt()(E): End Function
Function EnmsyTblFmt() As String()
Static Sy$(): If Si(Sy) = 0 Then Sy = NyQtp(EnmqssTblFmt)
EnmsyTblFmt = Sy
End Function
Function EnmvTblFmt(S) As eTblFmt: EnmvTblFmt = EnmvEnms(EnmsyTblFmt, S, "eTblFmt"): End Function
Function EnmsZer$(E As eZer):         EnmsZer = EnmsyZer()(E):                       End Function
Function EnmsyZer() As String()
Static Sy$(): If Si(Sy) = 0 Then EnmsyZer = NyQss(EnmqssZer)
EnmsyZer = Sy
End Function
Private Sub B_FmtDrs()
Dim D As Drs, Rds As eRds, Fmt As eTblFmt, Zer As eZer, Wdt%, NoIx As Boolean, ColnnAliR$, ColnnSum$
GoSub Z
Exit Sub
T1:
    D = sampDrs
    Wdt = 100
    GoSub Tst
Tst:
    Stop
    Rds = eRdsNo
    Wdt = 120
    Zer = eZerHid
    NoIx = False
    Fmt = eTblFmtTb
    ColnnAliR = ""
    Act = FmtDrs(D, Rds, Fmt, Zer, Wdt, ColnnAliR, ColnnSum, NoIx)
    Brw Act: Stop
    C
    Return
Z:
    Wdt = 120
    D = sampDrs1
    BrwAy FmtDrs(D, Rds, Fmt, Zer, Wdt, ColnnAliR, ColnnSum, NoIx)
    Return
End Sub
Function FmtDrs(D As Drs, _
Optional Rds As eRds, _
Optional Fmt As eTblFmt, _
Optional Zer As eZer, _
Optional Wdt% = 100, _
Optional ColnnAliR$, _
Optional ColnnSum$, _
Optional NoIx As Boolean _
) As String()
If Rds = eRdsYes Then
    FmtDrs = WFR_FmtRds(D, Fmt, Zer, Wdt, ColnnAliR, ColnnSum, NoIx)
Else
    FmtDrs = WFN_FmtNrm(D, Fmt, Zer, Wdt, ColnnAliR, NoIx)
End If
End Function

Private Function WFR_FmtRds(D As Drs, _
Fmt As eTblFmt, _
Zer As eZer, _
Wdt%, _
ColnnAliR$, _
ColnnSum$, _
NoIx As Boolean) As String()
Dim R As RdsDrs: R = RdsDrs(D, ColnnSum)
Dim H$(): H = W2FmtHdr(R, Fmt)
Dim T$(): T = WFN_FmtNrm(R.Drs, Fmt, Zer, Wdt, ColnnAliR, NoIx) ' ColnnSum$ and Tit$ are set to "" in calling !WFN_FmtNrm
WFR_FmtRds = SyAdd(H, T)
End Function
Private Function W2FmtHdr(R As RdsDrs, Fmt As eTblFmt) As String() ' Reduced format header
Dim A$(): If R.DupCol.Count > 0 Then A = FmtDi(R.DupCol, False, "Skip-Col See", Fmt)
Dim B$(): If R.CnstCol.Count > 0 Then B = FmtDi(R.CnstCol, True, "Cnst Val", Fmt)
Dim C$(): If R.FldSum.Count > 0 Then C = FmtDi(R.FldSum, False, "Col Sum", Fmt)
W2FmtHdr = LyDcyStrAp("", A, B, C)
End Function

Private Function WFN_FmtNrm(D As Drs, _
Fmt As eTblFmt, _
Zer As eZer, _
Wdt%, _
ColnnAliR$, _
NoIx As Boolean) As String()
'Stop '' Need to handle @ColnnSum
Dim Bdy$()
    Dim AlirCii$: AlirCii = CiiColnn(D.Fny, ColnnAliR)
    Dim Dy(): Dy = AyAyItm(D.Dy, D.Fny)
    Bdy = FmtDy(Dy, Fmt, Zer, Wdt, AlirCii, NoIx)

Dim Sepln$, Hdrln$, SeplnLas$
If Fmt = eTblFmtTb Then
    Sepln = Pop(Bdy)
    Hdrln = W1Hdrln(Pop(Bdy))
    If EleLas(Bdy) <> Sepln Then SeplnLas = Sepln
    WFN_FmtNrm = SyNB(Sepln, Hdrln, Bdy, SeplnLas)
Else
    Hdrln = Pop(Bdy)
    WFN_FmtNrm = SyAdd(LyUL(Hdrln), Bdy)
End If
End Function
Private Function W1Hdrln$(Hdrln$)
Dim O$: O = Hdrln
Dim J%: For J = 3 To InStr(2, Hdrln, "|") - 2
    Mid(O, J, 1) = "#"
Next
W1Hdrln = O
End Function
Function FmtDrsNmv(Lbl$, A As Drs) As String(): FmtDrsNmv = MsglyNmV(Lbl, FmtDrs(A)): End Function
