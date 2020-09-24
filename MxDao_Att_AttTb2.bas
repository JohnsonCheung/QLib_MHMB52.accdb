Attribute VB_Name = "MxDao_Att_AttTb2"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Att_AttTb2."
Private Sub B_CrtTb2AttNrmC()
DrpTb2AttC
CrtTb2AttNrmC
End Sub
Sub CrtTb2AttNrmC(): CrtTb2AttNrm CDb: End Sub
Sub CrtTb2AttNrm(D As Database)
D.TableDefs.Append WTdAtt
D.TableDefs.Append WTdAttd
End Sub
Private Function WTdAtt() As Dao.TableDef
Dim O As New Dao.TableDef
With O
    .Name = "Att"
    .Fields.Append FdId("AttId")
    .Fields.Append FdNNTxt("Attn")
    .Fields.Append FdAtt("Att")
    AddSk O, "Attn"
End With
Set WTdAtt = O
End Function
Private Function WTdAttd() As Dao.TableDef
Dim O As New Dao.TableDef
With O
    .Name = "Attd"
    .Fields.Append FdId("AttdId")
    .Fields.Append FdNNLng("AttId")
    .Fields.Append FdNNTxt("Fn")
    .Fields.Append FdNNDte("FfnTim")
    .Fields.Append FdNNLng("FfnLen")
    .Fields.Append FdNNDte("ImpTim")
    .Fields.Append FdMem("ImpFmFfn")
    AddSk O, "AttId Fn"
End With
Set WTdAttd = O
End Function

Sub CrtTb2AttSchm(D As Database): SchmCrt D, WSchmAtt: End Sub
Sub CrtTb2AttSchmC():             CrtTb2AttSchm CDb:   End Sub
Private Function WSchmAtt() As String()
Const A$ = "Tbl"
Const B1$ = "  Att * *n | Att"
Const B2$ = "  Attd * AttId Fn | FIlTim FfnLen"
Const C$ = "EleFld"
Const D$ = "  T22 FilTimSi22"
Const E$ = "  Att AttFn"
Const F$ = "  Nm  Attn"
WSchmAtt = SyAp(A, B1, B2, C, D, E, F)
End Function

Private Sub B_EnsTb2Att()
EnsTb2Att CDb
BrwDb CDb
End Sub
Sub EnsTb2AttC():             EnsTb2Att CDb:       End Sub
Sub EnsTb2Att(D As Database): SchmEns D, WSchmAtt: End Sub
Sub DrpTb2Att(D As Database): DrpTT D, "Att Attd": End Sub
Sub DrpTb2AttC():             DrpTb2Att CDb:       End Sub
