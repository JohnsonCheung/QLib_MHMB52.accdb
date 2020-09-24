Attribute VB_Name = "MxDao_Dbt_P_Stru"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Dbt_P_Stru."

Function StruFld(ParamArray Ap()) As Drs
Dim Dy(), S$, I, Av(), Ele$, LikFf$, LikFld$, J
Av = Ap
For Each I In Av
    S = I
    AsgT1r S, Ele, LikFf
    For Each J In SySs(LikFf)
        LikFld = J
        PushI Dy, Array(Ele, LikFld)
    Next
Next
StruFld = DrsFf("Ele FldLik", Dy)
End Function

Function StruInf(D As Database) As Dt
Dim T$, TT, Dy(), Des$, NRec&, Stru$
'For Each TT In TnyDb(D)
    T = TT
'    Des = Dbt_Des(D, T)
'    Stru = RmvA1T(Stru(D, T))
'    NRec = NRecDT(D, T)
    PushI Dy, Array(T, NRec, Des, Stru)
'Next
StruInf = DtFf("Tbl", "Tbl NRec Des", Dy)
End Function

Function StruC() As String():             StruC = Stru(CDb):                   End Function
Function Stru(D As Database) As String():  Stru = FmtT1ry(StruTny(D, Tny(D))): End Function
Function StruTny(D As Database, Tny$()) As String()
Dim I: For Each I In Itr(AySrtQ(Tny))
    PushI StruTny, StruT(D, I)
Next
End Function
Function StruRs$(A As Dao.Recordset)
Dim O$(), F As Dao.Field2
For Each F In A.Fields
    PushI O, StrFd(F)
Next
StruRs = JnCrLf(O)
End Function
Function StruTC$(T):                                 StruTC = StruT(CDb, T):          End Function
Function StruTTC(TT$) As String():                  StruTTC = StruTT(CDb, TT):        End Function
Function StruTT(D As Database, TT$):                 StruTT = StruTny(D, Tmy(TT)):    End Function
Function StruQtp(D As Database, Qtp$) As String():  StruQtp = StruTny(D, NyQtp(Qtp)): End Function
Function StruQtpC(Qtp$) As String():               StruQtpC = StruQtp(CDb, Qtp):      End Function
Function StruT$(D As Database, T)
Dim F$(), S$()
    F = FnyIf(D, T): If Si(F) = 0 Then StruT = "Tbn[" & T & "] not found in Db[" & D.Name & "]": Exit Function
    S = FnySk(D, T)
Dim Ss$
    If Si(S) > 0 Then
        Ss = " " & Tml(AmRplStar(S, T)) & " |"
    End If

Dim RR$
    Dim R$(): R = AyMinus(F, S)
    If Si(R) > 0 Then
        RR = " " & Tml(AmRplStar(R, T))
    End If
StruT = T & Ss & RR
End Function
