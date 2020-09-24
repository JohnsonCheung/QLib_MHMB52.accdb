Attribute VB_Name = "MxDao_Db_Op_CrtFb_CpyTbl"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Db_Op_CrtFb_CpyTbl."
Private Sub B_CrtFbCpyTblC()
Dim F$: F = FbTmp
CrtFbCpyTblC F
BrwFb F
Stop
End Sub
Sub CrtFbCpyTbl(Fb$, D As Database): CrtFbCpyTny Fb, Tny(D), D: End Sub
Sub CrtFbCpyTblC(Fb$):               CrtFbCpyTbl Fb, CDb:       End Sub

Sub CrtFbCpyOupC(Fb$):               CrtFbCpyOup Fb, CDb:          End Sub ' Create a new Fb by copy all @* from CDb
Sub CrtFbCpyOup(Fb$, D As Database): CrtFbCpyTny Fb, TnyOup(D), D: End Sub
Sub CrtFbCpyTny(Fb$, Tny$(), D As Database)
CrtFb Fb
Dim Q: For Each Q In Itr(WSqy(Fb, Tny, D))
'    Debug.Print Q
    D.Execute Q
Next
End Sub
Private Function WSqy(Fb$, Tny$(), D As Database) As String()
Const C$ = "Select * into [?] in '{Fb}' from [?]"
Dim Tp$: Tp = Rpl(C, "{Fb}", Fb)
Dim T: For Each T In Itr(Tny)
    Dim M$: M = RplQ(Tp, T)
        Dim F$: F = CmaFldNoAtt(D, T)
        If F <> "*" Then M = Rpl(M, "*", F)
    PushI WSqy, M
Next
End Function
