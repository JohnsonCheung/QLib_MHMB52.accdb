Attribute VB_Name = "MxDao_Dbt_Op_RseqFld"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Dbt_Op_RseqFld."
Public Const SampSpRseq = _
"*Flg Amt Qty *Key *Uom MovTy *Rate *Bch *Las *GL" & _
"|*Flg  IsAlert IsWithSku" & _
"|*Key  Sku PstMth PstDte RecTy" & _
"|*Rate BchRateUX RateTy" & _
"|*Bch  BchNo BchPermitDate BchPermit" & _
"|*Las  LasBchNo LasPermitDate LasPermit" & _
"|*GL   GLDocNo GLDocDte GLAsg GLDocTy GLLin GLPstKy GLPc GLAc GLBusA GLRef" & _
"|*Uom  Des StkUom Ac_U"

Private Sub B_FnySpRseq(): D FnySpRseq(SampSpRseq): End Sub
Function FnySpRseq(SpRseq$) As String()
Dim Ly$(): Ly = SplitVBar(SpRseq)
Dim NyLn1$(): NyLn1 = SySs(Ly(0))
Dim Di As Dictionary: Set Di = WDi(Ly)
Dim O$()
Dim T: For Each T In NyLn1
    If ChrFst(T) = "*" Then
        PushIAy O, SySs(Di(T))
    Else
        PushI O, T
    End If
Next
FnySpRseq = O
End Function
Private Function WDi(LySpRseq$()) As Dictionary
WChk LySpRseq
Set WDi = New Dictionary
Dim J%: For J = 1 To UB(LySpRseq)
    With BrkSpc(LySpRseq(J))
        WDi.Add .S1, .S2
    End With
Next
End Function
Private Sub WChk(LySpRseq$())
Const CSub$ = CMod & "WChk"
Dim Stary$(): Stary = AwPfx(SySs(LySpRseq(0)), "*")
Dim StaryDfn$()
    Dim J%: For J = 1 To UB(LySpRseq)
        PushI StaryDfn, Tm1(LySpRseq(J))
    Next
Dim StaryMis$(): StaryMis = SyMinus(Stary, StaryDfn)
Dim StaryExa$(): StaryExa = SyMinus(StaryDfn, Stary)
If Si(StaryMis) = 0 And Si(StaryExa) = 0 Then Exit Sub
Dim Msg$: Msg = FmtQQ("There are [?/?] missing/extra Star-Dfn", Si(StaryMis), Si(StaryExa))
Thw CSub, Msg, "Mis-Star Exa-Star LySpRseq", StaryMis, StaryExa, LySpRseq
End Sub

Sub RseqFld(D As Database, T, ByFny$())
Dim F, J%
For Each F In AyRseq(Fny(D, T), ByFny)
    J = J + 1
    D.TableDefs(T).Fields(F).OrdinalPosition = J
Next
End Sub

Sub RseqFldSp(D As Database, T, SpRseq$): RseqFld D, T, FnySpRseq(SpRseq): End Sub

Sub UpdFldSno(D As Database, T, FldSno$, FfGp$, Optional FfHypSfxOrdExtra$)
Dim Q$: Q = SqlSelFf(T, FldSno & " " & FfGp, , , FfGp & " " & FfHypSfxOrdExtra)
Dim R As Recordset: Set R = Rs(D, Q)
If NoRec(R) Then Exit Sub
Dim Seq&, DrLas(), DrCur(), N%
With R
    N = .Fields.Count - 1
    .MoveNext
    DrLas = DrRs(R)
    While Not .EOF
        DrCur = DrRsFstN(R, N)
        If Not IsEqAy(DrCur, DrLas) Then
            DrCur = DrLas
            Seq = 0
        End If
        Seq = Seq + 1
        .Edit
        R(FldSno).Value = Seq
        .Update
        .MoveNext
    Wend
End With
End Sub


Private Sub B_UpdFldSno()
Dim Db As Database, T$
Set Db = DbTmp
Runq Db, "Select * into [#A] from [T] order by Sku,PermitDate"
Runq Db, "Update [#A] set BchRateSeq=0, Rate=Round(Rate,0)"
UpdFldSno Db, T, "BchRateSeq", "Sku", "Sku Rate"
Stop
Drp Db, "#A"
End Sub
