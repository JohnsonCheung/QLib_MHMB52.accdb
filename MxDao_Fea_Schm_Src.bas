Attribute VB_Name = "MxDao_Fea_Schm_Src"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Feat_Schm_Src."
Const KwEle$ = "Ele"
Const KwEleFld$ = "EleFld"
Const KwFldDes$ = "PvFDesC"
Const KwKey$ = "Key"
Const KwTbl$ = "Tbl"
Const KwTblDes$ = "PvTDes"
Const KwTblFldDes$ = "TblFldDes"
Private S$()
Private Sub B_SchmSrc()
Dim Act As SchmSrc, Schm$(), Ept As SchmSrc
GoSub T1
Exit Sub
T1:
    Schm = SchmSamp(1)
    GoTo Tst
Tst:
    Act = SchmSrcSchm(Schm)

    Stop
    Return
End Sub
Function SchmSrcSchm(Schm$()) As SchmSrc
S = Schm
With SchmSrcSchm
    .Ele = WEle(WLLny(KwEle))
    .EleFld = WEleFld(WLLny(KwEleFld))
    .PvFDesC = WFldDes(WLLny(KwFldDes))
    .Key = WKey(WLLny(KwKey))
    .Tbl = WTbl(WLLny(KwTbl))
    .PvTDes = WTblDes(WLLny(KwTblDes))
    .TblFldDes = WTblfldDes(WLLny(KwTblFldDes))
End With
Stop
End Function
Private Function WLLny(Key$) As LLn()
Stop
End Function
Private Function WEle(L() As LLn) As SmsEle()
Dim M As LLn
Dim J%: For J = 0 To LLnUB(L)
    M = L(J)
    PushSmsEle WEle, SmsEle(M.Lno, Tm1(M.Ln), RmvA1T(M.Ln))
Next
End Function

Private Function WEleFld(L() As LLn) As SmsEleFld()
Dim M As LLn
Dim J%: For J = 0 To LLnUB(L)
    M = L(J)
    PushSmsEleFld WEleFld, SmsEleFld(M.Lno, Tm1(M.Ln), SySs(RmvA1T(M.Ln)))
Next
End Function
Private Function WFldDes(L() As LLn) As SmsFldDes()

End Function
Private Function WKey(L() As LLn) As SmsEle()

End Function
Private Function WTbl(L() As LLn) As SmsTbl()
Dim M As LLn
Dim J%: For J = 0 To LLnUB(L)
    M = L(J)
    Dim Tbn$, Rst$: AsgT1r M.Ln, Tbn, Rst
    Dim Fny$(), FnySk$(): WTblAsg Tbn, Rst, Fny, FnySk
    PushSmsTbl WTbl, SmsTbl(M.Lno, Tbn, Fny, FnySk)
Next
Stop
End Function
Private Sub WTblAsg(Tbn$, Rst$, OFny$(), OSkFny$())
Dim R$: R = Replace(Rst, "*", Tbn)
Dim S As S12: S = Brk2(R, "|")
OSkFny = SySs(S.S1)
OFny = SyAdd(OSkFny, SySs(S.S2))
End Sub
Private Function WTblDes(L() As LLn) As SmsTblDes()

End Function
Private Function WTblfldDes(L() As LLn) As SmsTblFldDes()

End Function
