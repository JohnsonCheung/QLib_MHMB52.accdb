Attribute VB_Name = "MxDao_Dbt_OpDrp"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Dbt_OpDrp."
Sub DrpC(T): Drp CDb, T: End Sub
Sub Drp(D As Database, T)
If HasT(D, T) Then D.Execute "Drop Table [" & T & "]"
End Sub
Sub DrpQtp2(D As Database, Qtp1$, Rst$): DrpTny D, NyQtp2(Qtp1, Rst): End Sub
Sub DrpQtp2C(Qtp1$, Rst$):               DrpQtp2 CDb, Qtp1, Rst:      End Sub
Sub DrpQtpC(Qtp$):                       DrpQtp CDb, Qtp:             End Sub
Sub DrpQtp(D As Database, Qtp$):         DrpTny D, NyQtp(Qtp):        End Sub
Sub DrpTmpC():                           DrpTmp CDb:                  End Sub
Sub DrpTmp(D As Database):               DrpTny D, TnyTmp(D):         End Sub
Sub DrpHshC():                           DrpHsh CDb:                  End Sub
Sub DrpHsh(D As Database):               DrpTny D, TnyHsh(D):         End Sub
Sub DrpTny(D As Database, Tny)
Dim T: For Each T In Itr(Tny)
    Drp D, T
Next
End Sub
Sub DrpTnyC(Tny):                             DrpTny CDb, Tny:                      End Sub
Sub DrpTTC(TT$):                              DrpTT CDb, TT:                        End Sub
Sub DrpTT(D As Database, TT$):                DrpTny D, SySs(TT):                   End Sub
Sub DrpApC(ParamArray TbnAp()):               Dim Av(): Av = TbnAp: DrpTny CDb, Av: End Sub
Sub DrpAp(D As Database, ParamArray TbnAp()): Dim Av(): Av = TbnAp: DrpTny D, Av:   End Sub
Sub DrpPfxxC(Pfxx$):                          DrpPfxx CDb, Pfxx:                    End Sub
Sub DrpPfxx(D As Database, Pfxx$):            DrpTny D, AwPfxx(Tny(D), Pfxx):       End Sub
