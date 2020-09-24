Attribute VB_Name = "MxDao_Fea_Lim_LnkImp"
Option Compare Text
Const CMod$ = "MxDao_Lim_LnkImp."
Option Explicit
Private WDbLim As Database, WLimn$
Sub LnkImpzLimnC(FbLim$, Limn$): LnkImpzLimn CDb, FbLim, Limn: End Sub
Sub LnkImpzLimn(D As Database, FbLim$, Limn$)
Set WDbLim = Db(FbLim)
WLimn = Limn
WChk
LnkTLnkTbl D, WTLnkTbl
RunqSqy D, WSqyImp
End Sub
Private Sub WChk()

End Sub
Private Function WTLnkTbl() As TLnkTbl
WTLnkTbl.Fb = TLnkTblFby
WTLnkTbl.Fx = TLnkTblFxy
End Function
Private Function WFbny() As String()
Stop
End Function
Private Function WFxny() As String()
Stop
End Function
Private Function TLnkTblFby() As TLnkTblFb()
Dim Fbn: For Each Fbn In Itr(WFbny)
    PushTLnkTblFb TLnkTblFby, WTLnkTblFb(Fbn)
Next
End Function
Private Function TLnkTblFxy() As TLnkTblFx()
Dim Fxn: For Each Fxn In Itr(WFxny)
    PushTLnkTblFx TLnkTblFxy, WTLnkTblFx(Fxn)
Next
End Function
Private Function WTLnkTblFb(Fbn) As TLnkTblFb

End Function
Private Function WTLnkTblFx(Fxn) As TLnkTblFx
Dim Wsny$()
Dim Tny$()
Dim TnyAs$()
    WAsg3ColTLnkTblFx Fxn, Wsny, Tny, TnyAs
Dim Fx$: Fx = WFxFxn(Fxn)
Dim I%: For I = 0 To UB(Wsny)
    With WTLnkTblFx
        .Fx = Fx
        .T = Tny(I)
        .TbnAs = TnyAs(I)
        .WsnFx = Wsny(I)
    End With
Next
End Function
Private Function WFxFxn$(Fxn)
Stop
End Function
Private Sub WAsg3ColTLnkTblFx(Fxn, OWsny$(), OTny$(), OTnyAs$())
Dim R As Dao.Recordset
Const C$ = "Select Wsn,Tbn,TbnAs from LimFxWs x inner join LimFx a a.FxId=x.FxId where a.Fxn='?' and a.Limn='?'"
Set R = RsQQC(C, Fxn, WLimn)
With R
    While Not .EOF
        PushI OWsny, !Wsn
        PushI OTny, !Tbn
        PushI OTnyAs, !TbnAs
        .MoveNext
    Wend
End With
End Sub
Private Function WSqyImp() As String()
WSqyImp = SyAddAp(WSqyCrt, WSqyPk, WSqySk, WSqyKey)
End Function
Private Function WSqyCrt() As String()

End Function
Private Function WSqyPk() As String()

End Function
Private Function WSqySk() As String()

End Function
Private Function WSqyKey() As String()

End Function
