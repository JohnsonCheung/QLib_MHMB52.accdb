Attribute VB_Name = "MxDao_Att_AttOp_IExpAtt"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Att_Do_AttIExp."
Sub ExpAtt(D As Database, Attn$, Attfn$, Ffn$)
Const CSub$ = CMod & "ExpAtt"
Sts FmtQQ("ExpAtt Attn[?] AttFn[?] Ffn[?] Db[?]", Attn, Attfn, Ffn, D.Name)
If Ext(Attfn) <> Ext(Ffn) Then Thw CSub, "Ext of Attn and Ffn should be same", "Attn AttFn Ffn", Attn, Attfn, Ffn
ChkFfnNExi Ffn, "ExpAtt", "Target File to be overwritten by Att[" & Attn & "] AttFn[" & Attfn & "]"
Dim R As Dao.Recordset: Set R = RsQQ(D, "Select Att from Att where Attn='?'", Attn)
    If NoRec(R) Then Thw CSub, "No Attn[?] in Tb-Att", "Attn", Attn
Dim R2 As Dao.Recordset2: Set R2 = Rs2Att(R, Attfn)
    If IsNothing(R2) Then Thw CSub, "In Tb-Att->Att, there is no Rs2-record with !FileName=AttFn_", "Attn AttFn", Attn, Attfn
CvFd2(R2.Fields("FileData")).SaveToFile Ffn
R2.Close
R.Close
End Sub
Sub ExpAttC(Attn$, Attfn$, FfnTo$): ExpAtt CDb, Attn, Attfn, FfnTo: End Sub
Sub ImpAtt(Ffn$, D As Database, Attn$, Optional AttFn_0$)
Const CSub$ = CMod & "ImpAtt"
ChkPm:
    ChkFfnExi Ffn, CSub, "Imp-To-Att-Ffn"
    If Len(Attn) > 255 Then Thw CSub, "Attn-Len cannot >255", "Attn-Len Attn", Len(Attn), Attn
    Dim Attfn$: Attfn = StrDft(AttFn_0, Fn(Ffn))
    If Ext(Ffn) <> Ext(Attfn) Then ThwPm CSub, "Ext of Ffn & AttFn should be same", "Ffn AttFn", Ffn, Attfn

Dim R As Dao.Recordset: Set R = W1EnsTbAttRec(D, Attn) '<=== Tb-Att may add a new record
Dim R2 As Dao.Recordset2: Set R2 = Rs2Att(R, Attfn)
If Not IsNothing(R2) Then
    If W1IsAttfSam(D, Ffn, CLng(R!AttId), Attfn) Then
        Debug.Print FmtQQ("?: Ffn is same, no import.  Attn[?] Attf[?] Ffn[?] FfnTim[?] FfnLen[?]", CSub, Attn, Attfn, Ffn, FileDateTime(Ffn), FileLen(Ffn))
        Exit Sub
    End If
    R.Edit
    R2.Delete  '<== always delete att no matter newer or older
    R.Update
End If
W1_Imp D, Attn, Attfn, Ffn            '<== Imp
W1UpdTbAttd D, Attn, Attfn, Ffn  '<== Upd Tb-Attd
Debug.Print FmtQQ("?: Ffn IMPORTED.  Attn[?] Attf[?] Ffn[?] FfnTim[?] FfnLen[?]", CSub, Attn, Attfn, Ffn, FileDateTime(Ffn), FileLen(Ffn))
End Sub
Sub ImpAttC(Ffn$, Attn$, Optional AttFn_0$): ImpAtt Ffn, CDb, Attn, AttFn_0: End Sub
Private Sub B_ExpAtt()
Dim T$, D As Database
T = FxTmp
ExpAtt CDb, "Tp", "MHMB52Io.Tp.xlsx", T
Debug.Assert HasFfn(T)
Kill T
End Sub
Private Sub B_ImpAttC()
Dim Ffn$
GoSub Z1
Exit Sub
ZZ:
    Ffn = FtTmp
    WrtStr "sdfdf", Ffn
    ImpAttC Ffn, "AA", "XX.txt"
    Return
Z1:
    ImpAttC MHO.MHOMB52.TpMB52, "Tp", MHO.MHOMB52.TpMB52
    Return
Stop
Const Fx$ = "C:\Users\Public\Logistic\StockHolding8\WorkingDir\Templates\Stock Holding Template.xlsx"
End Sub
Private Sub W1_Imp(D As Database, Attn$, Attfn$, Ffn$)
Dim R As Dao.Recordset: Set R = RsAtt(D, Attn)
Dim Rs2 As Dao.Recordset2: Set Rs2 = R.Fields("Att").Value
Dim Fd2 As Dao.Field2: Set Fd2 = Rs2!FileData
R.Edit
Rs2.AddNew
    If Fn(Ffn) <> Attfn Then
        Dim A$: A = Pth(Ffn) & Attfn
        CpyFfn Ffn, A
        Fd2.LoadFromFile A      '<==
        Rs2.Update
        DltFfn A
    Else
        Fd2.LoadFromFile Ffn        '<==
        Rs2.Update
    End If
R.Update
End Sub
Private Function W1EnsTbAttRec(D As Database, Attn$) As Dao.Recordset   ' Ensure a record of @Attn in tb-Att
Dim Q$: Q = SqlSelStarFeq("Att", "Attn", Attn)
Dim O As Dao.Recordset: Set O = Rs(D, Q)
If HasRec(O) Then
    Set W1EnsTbAttRec = O
Else
    D.Execute FmtQQ("Insert into Att (Attn) values ('?')", Attn)
    Set W1EnsTbAttRec = Rs(D, Q)
End If
End Function
Private Function W1IupTbAttd(D As Database, Attn$, Attfn$) As Dao.Recordset: Set W1IupTbAttd = RsSkvapEdt(D, "Attd", AttId(D, Attn), Attfn): End Function
Private Sub W1UpdTbAttd(D As Database, Attn$, Attfn$, FfnFm$)
With W1IupTbAttd(D, Attn, Attfn)
!FfnTim = FileDateTime(FfnFm)
!FfnLen = FileLen(FfnFm)
!ImpFmFfn = FfnFm
!ImpTim = Now()
.Update
End With
End Sub
Private Function W1IsAttfSam(D As Database, Ffn$, AttId&, Attfn$) As Boolean
Dim R As Dao.Recordset: Set R = Rs(D, SqlSelStarFfeq("Attd", "AttId Fn", Array(AttId, Attfn)))
If R!FfnTim <> FileDateTime(Ffn) Then Exit Function
If R!FfnLen <> FileLen(Ffn) Then Exit Function
W1IsAttfSam = True
End Function
