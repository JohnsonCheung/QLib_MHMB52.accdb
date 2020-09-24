Attribute VB_Name = "MxVb_Git"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Git."
Const Fgit$ = "C:\Program Files\Git\Cmd\git.Exe"
Const FgitQ$ = vbQuoDbl & Fgit & vbQuoDbl
Sub Cmit(Optional Msg$ = "commit", Optional ReInit As Boolean): ShellMaxCdl "Cmit", WCdlFcmdCmit(Msg, ReInit): End Sub
Private Function WCdlFcmdCmit$(Optional Msg$ = "Commit", Optional ReInit As Boolean)
XX
    Dim P$:     P = PthSrcPC
    X "Cd """ & P & """"
    X FmtQQ("Set FexeGit=?", FgitQ)
    If ReInit Then X "Rd .git /s/q"
    X "%FexeGit% init"       'If already init, it will do nothing
    X "%FexeGit% add -A"
    X FmtQQ("%FexeGit% commit -m ""?""", Msg)
    X "Pause"
WCdlFcmdCmit = JnCrLf(XX)
'Debug.Print WCdlFcmdCmit: Stop
End Function

Sub GitPush():                ShellMaxCdl "Push", WCdlFcmdPush: End Sub
Private Sub B_WCdlFcmdPush(): BrwStr WCdlFcmdPush:              End Sub
Private Function WCdlFcmdPush$()
XX
    Dim P$: P = PthSrcPC
    X FmtQQ("Cd ""?""", P)
    X FmtQQ("? push -u https://johnsoncheung@github.com/johnsoncheung/?.git master", FgitQ, WPjnPthSrc(P))
    X "Pause ....."
WCdlFcmdPush = JnCrLf(XX)
'Debug.Print WCdlFcmdPush: Stop
End Function

Sub ChkHasInternet(): RaiseTrue Not HasInternet, "No internet": End Sub
Function HasInternet() As Boolean
Stop
End Function

Private Sub B_WPjnPthSrc()
Dim PthSrc$
GoSub T1
Exit Sub
T1:
    PthSrc = "C:\Users\user\Documents\Projects\Vba\QLib\.{QLib_StockHolding8}.accdb\.src\"
    Ept = "{QLib_StockHolding8}"
    GoTo Tst
Tst:
    Act = WPjnPthSrc(PthSrc)
    C
    Return
End Sub
Private Function WPjnPthSrc$(PthSrc$)
'Sample-PthSrc-value: C:\Users\user\Documents\Projects\Vba\QLib\.{QLib_StockHolding8}.accdb\.src
'Ret                : {QLIb_StockHolding8}
Const CSub$ = CMod & "WPjnPthSrc"
Dim P$: P = PthRmvSfx(PthSrc)
If Fdr(P) <> ".src" Then Thw CSub, "Not a source path: Fdr should be [.src]", "PthSrc Fdr", PthSrc, Fdr(PthSrc)
Dim B$: B = Fn(Fdr(PthPar(P))) ' SHould be .{QLib_StockHolding8}
Dim C$: C = ChrFst(B)  ' Should be .
If ChrFst(C) <> "." Then Thw CSub, "Fst Chr of X:Fn(Fdr(PthPar(@PthSrc)))) should be [.]", "PthSrc X", PthSrc, C
WPjnPthSrc = RmvFst(B)
End Function
