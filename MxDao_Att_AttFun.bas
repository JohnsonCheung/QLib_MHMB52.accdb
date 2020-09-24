Attribute VB_Name = "MxDao_Att_AttFun"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Att_AttFun."

Function TimAttf(D As Database, Attn$, Attfn$) As Date: TimAttf = WRsAttfnTimLen(D, Attn, Attfn).Fields("FfnTim").Value: End Function
Function LenAttf&(D As Database, Attn$, Attfn$):        LenAttf = WRsAttfnTimLen(D, Attn, Attfn).Fields("FfnLen").Value: End Function
Private Function WRsAttfnTimLen(D As Database, Attn$, Attf$) As Dao.Recordset
Const C$ = "Select FfnTim,FfnLen from Attd where Fn='?' and Att_AttId=?"
Set WRsAttfnTimLen = Rs(D, FmtQQ(C, Attf, AttId(D, Attn)))
End Function

Private Sub B_AttfnayAttn():                                            D AttfnayAttn(CDb, "Att"):                                                               End Sub
Function Attfnay(D As Database) As String():                  Attfnay = DcStrTF(D, "Attd.Fn"):                                                                   End Function
Function AttfnayC() As String():                             AttfnayC = Attfnay(CDb):                                                                            End Function
Function AttfnayAttnC(Attn$) As String():                AttfnayAttnC = AttfnayAttn(CDb, Attn):                                                                  End Function
Function AttfnayAttn(D As Database, Attn$) As String():   AttfnayAttn = DcStrQ(D, SqlSelFldWhFeq("Attd", "Fn", "AttId", AttId(D, Attn))):                        End Function
Function AttfnayColon(D As Database) As String():        AttfnayColon = DcStrQ(D, "Select Attn & ':' & Fn from Att x inner join Attd a on a.Att_AttId=x.AttId"): End Function
Function AttfnaycolonC() As String():                   AttfnaycolonC = AttfnayColon(CDb):                                                                       End Function
Function Attn&(D As Database, AttId&):                           Attn = ValQ(D, "Select Attn where Att_AttId=" & AttId):                                         End Function

Function AttId&(D As Database, Attn$):      AttId = RsAtt(D, Attn).Fields("AttId"): End Function
Function Attny(D As Database) As String():  Attny = DcStrTF(D, "Att.Attn"):         End Function
Function AttnyC() As String():             AttnyC = Attny(CDb):                     End Function

Function FnyFd2C() As String(): FnyFd2C = FnyFd2(CDb): End Function
Function FnyFd2(D As Database) As String()
Const Q = "Select Att from Att where Attn='*Dft'"
Dim R As Dao.Recordset: Set R = Rs(D, Q)
If NoRec(R) Then
    D.Execute "Insert into Att (Attn) values ('*Dft')"
End If
FnyFd2 = FnyRs(Rs(D, Q).Fields("Att").Value)
End Function
Function RsAtt(D As Database, Attn$) As Dao.Recordset: Set RsAtt = RsTFeq(D, "Att.Attn", Attn): End Function

Function NAttFnzAttdC%(Attn$):               NAttFnzAttdC = NAttFnzAttd(CDb, Attn):                             End Function
Function NAttFnzAttd%(D As Database, Attn$):  NAttFnzAttd = NRecT(D, "Attd", BeprFeq("AttId", AttId(D, Attn))): End Function

Function DbTmpAtt() As Database
'Ret: a tmp db with tb-Att&Attd @@
Dim O As Database: Set O = DbTmp
EnsTb2Att O
Set DbTmpAtt = O
End Function
Function HasAttFnzTbAttd(D As Database, Attn$, Attfn$) As Boolean
Dim Q$: Q = FmtQQ("Select * from Attd x inner join Att a on x.Att_AttId=a.AttId where Fn='?' and Attn='?'", Attfn, Attn)
HasAttFnzTbAttd = HasRecQ(D, Q)
End Function
Function HasAttFnzTbAtt(D As Database, Attn$, Attfn$) As Boolean
Dim Q$: Q = FmtQQ("Select * from Att where Attn='?'", Attn)
Dim R As Dao.Recordset: Set R = Rs(D, Q)
If NoRec(R) Then Exit Function
HasAttFnzTbAtt = Not IsNothing(Rs2Att(R, Attfn))
End Function

Function Rs2Att(RsAtt As Dao.Recordset, Attf$) As Dao.Recordset2
Const CSub$ = CMod & "Rs2Att"
If NoRec(RsAtt) Then ThwPm CSub, "Given [RsAtt] should have a record"
Dim R2 As Dao.Recordset2: Set R2 = RsAtt.Fields("Att").Value
With R2
    While Not .EOF
        .MoveFirst
        While Not .EOF
            If !FileName = Attf Then Set Rs2Att = R2: Exit Function
            .MoveNext
        Wend
    Wend
    .Close
End With
End Function
