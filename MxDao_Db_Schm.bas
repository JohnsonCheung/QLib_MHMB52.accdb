Attribute VB_Name = "MxDao_Db_Schm"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Db_Schm."
Private S As SchmSrc
Type SchmDta
    Fnyy() As Variant
    TnyAll() As String
    FnyAll() As String
    EleAll() As String
    Di_TFdot_EleStr As Dictionary
    Di_Fld_Ele As Dictionary
    TnyPk() As String
    TFnyySk() As TbFny
End Type
Private D As SchmDta
Private Sub B_SchmCrt()
Dim D As Database, Schm$()
GoSub T1
Exit Sub
T1:
    Set D = DbTmp
    Schm = SchmSamp(1)
    GoTo Tst
Tst:
    SchmCrt D, Schm
    Return
End Sub

Private Sub B_SchmEnsC()
Dim Schm$()
YY:
    Schm = SchmSamp(1)
    GoTo Tst
Tst:
    SchmEnsC Schm
    Stop
    Return
End Sub

Sub SchmEnsC(Schm$()): SchmEns CDb, Schm: End Sub
Sub SchmEns(D As Database, Schm$())

End Sub
Sub SchmCrtC(Schm$()): SchmCrt CDb, Schm: End Sub
Sub SchmCrt(Db As Database, Schm$())
Const CSub$ = CMod & "SchmCrt"
Dim S As SchmSrc
S = SchmSrcSchm(Schm)
D = WDta_4
ChkEry Schm_Er(S, D), CSub
AppTdy Db, WTdy_2
RunqSqy Db, WSqyPk_3
RunqSqy Db, WSqySk
RunqSqy Db, WSqyIdx
RunqSqy Db, WSqyFk
SetPvDesDi Db, WDiTblDes
SetFDes Db, WDiFldDes
End Sub
Private Function WTdy_2() As Dao.TableDef()
Dim J%: For J = 0 To UB(D.TnyAll)
    With D
        Dim T$: T = .TnyAll(J)
        Dim Fny$(): Fny = .Fnyy(J)
    End With
    PushObj WTdy_2, W2_Td(T, Fny)
Next
End Function
Private Function W2_Td(T$, Fny$()) As Dao.TableDef
Dim Fdy() As Dao.Field
    Dim F: For Each F In Fny
        PushObj Fdy, FdEleStr(F, D.Di_TFdot_EleStr(T & "." & F))
    Next
Set W2_Td = TdFdy(T, Fdy)
End Function
Private Function WSqyPk_3() As String()
Dim TbnPk: For Each TbnPk In W3_TnyPk
    PushI WSqyPk_3, SqlCrtPk(TbnPk)
Next
End Function
Private Function W3_TnyPk() As String()
Dim J%: For J = 0 To UB(D.TnyAll)
    Dim Tbn$: Tbn = D.TnyAll(J)
    If HasEle(D.Fnyy(J), D.TnyAll(J) & "Id") Then PushI W3_TnyPk, Tbn
Next
End Function
Private Function WSqySk() As String()
Dim TF() As TbFny: TF = D.TFnyySk
Dim J%: For J = 0 To UbTbFny(TF)
    PushI WSqySk, SqlCrtSk(TF(J))
Next
End Function
Private Function WSqyIdx() As String()

End Function
Private Function WSqyFk() As String()

End Function
Private Function WDiTblDes() As Dictionary
End Function

Private Function WDiFldDes() As Dictionary
End Function

Private Function WDta_4() As SchmDta
With WDta_4
Set .Di_TFdot_EleStr = W4_Di_TFdot_EleStr
Set .Di_Fld_Ele = W4_Di_Fld_Ele
.TnyAll = W4_TnyAll
.FnyAll = W4_FnyAll
.EleAll = W4_EleAll
Stop '.Fnyy = W4_Fnyy(Tny,FssAy)
.TnyPk = W4_TnyPk
End With
End Function
Private Function W4_Di_TFdot_EleStr() As Dictionary
Dim AllFny$(): 'AllFny = AwDis(AyAyy(T_Fny))
Dim FqE As Dictionary: 'Set FqE = W4_Di_Fld_Ele(AllFny, EF_E, EF_FldLiky)
Dim EqEs As Dictionary: ' Set EqEs = DiAy12(E_E, E_EleStr)
Dim FqEs As Dictionary: Set FqEs = ChainDi(FqE, EqEs)
Stop 'Set Di_TFdot_EleStr = DiWhKy(FqEs, AllFny)
End Function
Private Function W4_Di_Fld_Ele() As Dictionary
Set W4_Di_Fld_Ele = New Dictionary
Dim Fld: For Each Fld In D.FnyAll
    Stop
    Dim Ix%: 'Ix = IxLikssAy(F, EF_FldLiky)
    Dim Ele$: 'Ele = D.EF_E(Ix)
    W4_Di_Fld_Ele.Add Fld, Ele
Next
End Function

Private Function W4_Fnyy(Tny$(), FssAy$()) As Variant()
Dim Fss, J%: For Each Fss In Itr(FssAy)
    PushI W4_Fnyy, SySs(Replace(Fss, "*", Tny(J)))
Next
End Function

Private Function W4_TnyAll() As String()
Dim J%: For J = 0 To UbSmsTbl(S.Tbl)
    PushS W4_TnyAll, S.Tbl(J).Tbn
Next
End Function
Private Function W4_FnyAll() As String()
End Function
Private Function W4_EleAll() As String()
End Function
Private Function W4_TnyPk() As String()

End Function
Private Function W4_TFnyySk() As TbFny()

End Function
