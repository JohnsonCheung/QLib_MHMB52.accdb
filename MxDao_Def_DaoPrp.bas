Attribute VB_Name = "MxDao_Def_DaoPrp"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Def_Prp."
Sub DltPv(Prps As Dao.Properties, P)
If HasPrp(Prps, P) Then
    Prps.Delete P
End If
End Sub
Sub DltPvDb(D As Database, P):                                   DltPv D.Properties, P:                          End Sub
Sub DltPvF(D As Database, T, F, P$):                             DltPv D.TableDefs(T).Fields(F).Properties, P:   End Sub
Sub DltPvFC(T, F$, P$):                                          DltPvF CDb, T, F, P:                            End Sub
Sub DltPvFDes(D As Database, T, F):                              DltPvF D, T, F, C_Des:                          End Sub
Sub DltPvFDesC(T, F):                                            DltPvFDes CDb, T, F:                            End Sub
Sub DltPvQryFld(D As Database, Q, F, P):                         DltPv D.QueryDefs(Q).Fields(F).Properties, P:   End Sub
Sub DltPvT(D As Database, T, P):                                 DltPv D.TableDefs(T).Properties, P:             End Sub
Sub DltPvTDes(D As Database, T):                                 DltPvT D, T, C_Des:                             End Sub
Sub DltPvTDesC(T):                                               DltPvTDes CDb, T:                               End Sub
Function HasPrp(Prps As Dao.Properties, P) As Boolean:  HasPrp = HasItn(Prps, P):                                End Function
Function HasPrpF(D As Database, T, F, P$) As Boolean:  HasPrpF = HasItn(D.TableDefs(T).Fields(F).Properties, P): End Function
Function HasPrpT(D As Database, T, P$) As Boolean:     HasPrpT = HasPrp(D.TableDefs(T).Properties, P):           End Function
Function PrpnyFd(A As Dao.Field) As String():          PrpnyFd = Itn(A.Properties):                              End Function
Function PvF(D As Database, T, F, P$)
If Not HasPrpF(D, T, F, P) Then Exit Function
PvF = D.TableDefs(T).Fields(F).Properties(P).Value
End Function
Function PvFDes$(D As Database, T, F):  PvFDes = PvF(D, T, F, C_Des): End Function
Function PvFDesC$(T, F):               PvFDesC = PvFDes(CDb, T, F):   End Function
Function PvT(D As Database, T, P$)
If Not HasPrpT(D, T, P) Then Exit Function
PvT = D.TableDefs(T).Properties(P).Value
End Function
Function PvTC(T, P$):                      PvTC = PvT(CDb, T, P):               End Function
Function PvTDes$(D As Database, T):      PvTDes = PvT(D, T, "Description"):     End Function
Function PvTDesC$(T):                   PvTDesC = PvTDes(CDb, T):               End Function
Function Pvln$(P As Dao.Property):         Pvln = P.Name & "=" & StrV(Objv(P)): End Function
Function PvlnAdo$(P As ADODB.Property): PvlnAdo = P.Name & "=" & StrV(Objv(P)): End Function
Function Pvlny(P As Dao.Properties) As String()
Dim I As Dao.Property: For Each I In P
    PushI Pvlny, Pvln(I)
Next
End Function
Function PvlnyAdo(P As ADODB.Properties) As String()
Dim I As ADODB.Property: For Each I In P
    PushI PvlnyAdo, PvlnAdo(I)
Next
End Function
Function PvlnyCDb() As String():               PvlnyCDb = PvlnyDb(CDb):               End Function
Function PvlnyDb(D As Database) As String():    PvlnyDb = Pvlny(D.Properties):        End Function
Function PvlnyFd(F As Dao.Field) As String():   PvlnyFd = Pvlny(F.Properties):        End Function
Function PvlnyT(D As Database, T) As String():   PvlnyT = Pvlny(Td(D, T).Properties): End Function
Sub SetPvF(D As Database, T, F$, P$, V)
Dim Td As Dao.TableDef: Set Td = D.TableDefs(T)
Dim Fd As Dao.Field: Set Fd = Td.Fields(F)
If WW_IsOkSetPv(Fd.Properties, P, V) Then Exit Sub
Fd.Properties.Append Fd.CreateProperty(P, Daoty(V), V)
End Sub
Sub SetPvFDes(D As Database, T, F$, Des$): SetPvF D, T, F, C_Des, Des: End Sub
Sub SetPvQryFld(D As Database, Q, F, P, V)
Dim Qd As QueryDef: Set Qd = CurrentDb.QueryDefs(Q)
Dim Fds As Dao.Fields: Set Fds = Qd.Fields
Dim Fd As Dao.Field: Set Fd = Fds(F)
If HasItn(F.Properties, P) Then
    Fd.Properties(P).Value = V
Else
    Fd.Properties.Append Fd.CreateProperty(P, Daoty(V), V)
End If
End Sub
Sub SetPvT(D As Database, T, P$, V)
Dim Td As Dao.TableDef: Set Td = D.TableDefs(T)
If WW_IsOkSetPv(Td.Properties, P, V) Then Exit Sub
Td.Properties.Append Td.CreateProperty(P, Daoty(V), V)
End Sub
Sub SetPvTC(T, P$, V):                 SetPvT CDb, T, P, V:             End Sub
Sub SetPvTDes(D As Database, T, Des$): SetPvT D, T, "Description", Des: End Sub
Sub SetPvTDesC(T, Optional Des$):      SetPvTDes CDb, T, Des:           End Sub
Private Function WW_IsOkSetPv(Ps As Dao.Properties, P$, V) As Boolean
If HasPrp(Ps, P) Then
    Ps(P).Value = V
    WW_IsOkSetPv = True
End If
End Function
