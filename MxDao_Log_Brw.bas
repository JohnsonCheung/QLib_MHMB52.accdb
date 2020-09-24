Attribute VB_Name = "MxDao_Log_Brw"
Option Compare Text
Const CMod$ = "MxDao_Log_Brw."
Option Explicit

Sub BrwSess(Optional SessId&)
Dim Ly$()
Dim Log&(): Log = LogIdy(SessId)
Stop 'BrwAy LySess(SessId)
End Sub
Function LogIdy(SessId&) As Long()
LogIdy = DcLngQ(DbLog, FmtQQ("select LogId from Log where SessId=? order by LogId", SessId))
End Function

Sub LisSess(Optional Sep$ = " ", Optional Top% = 50)
Dim R As Dao.Recordset: Set R = DbLog.OpenRecordset(FmtQQ("Select Top ? * from sess order by Sess desc", Top))
Dim Ly$(): Ly = LyJnRs(R, Sep)
D Ly
End Sub

Function NLogSess%(SessId&)
NLogSess = ValQ(DbLog, "Select Count(*) from Log where SessId=" & SessId)
End Function

Sub BrwPthLog(): BrwPth PthLog: End Sub
Sub BrwFbLog():  BrwFb FbLog:   End Sub

Sub BrwLog(Optional Top% = 50): BrwAy W0_LyLog(Top): End Sub

Sub LisLog(Optional Sep$ = " ", Optional Top% = 50)
Dim R As Dao.Recordset: Set R = DbLog.OpenRecordset(FmtQQ("Select Top ? x.*,Fun,MsgTxt from Log x left join Msg a on x.Msg=a.Msg order by Sess desc,Log", Top))
Dim Ly$(): Ly = LyJnRs(R, Sep)
D Ly
End Sub

Private Function W0_LyLog(Top%) As String()
Stop
End Function

Function LogLsy(LogId&) As Variant()
Dim Q$
Q = FmtQQ("Select Lines from LogV where Log = ? order by LogV", LogId)
'LogLsy = RsAy(L.OpenRecordset(Q))
End Function

Sub LogLis(Optional Sep$ = " ", Optional Top% = 50)
LisLog Sep, Top
End Sub

Function LogLy1(A&) As String()
Dim Fun$, MsgTxt$, StrDte$, Sess&, Sfx$
W00_Asg A, Sess, StrDte, Fun, MsgTxt
Sfx = FmtQQ(" @? Sess(?) Log(?)", StrDte, Sess, A)
'WLyLog = Fmsg(Fun & Sfx, MsgTxt, LogLsy(A))
Stop '
End Function
Private Sub W00_Asg(A&, OSess&, OTimStr_Dte$, OFun$, OMsgTxt$)
Dim Q$
Q = FmtQQ("select Fun,MsgTxt,Sess,x.CrtTim from Log x inner join Msg a on x.Msg=a.Msg where Log=?", A)
Dim D As Date
AsgRs DbLog.OpenRecordset(Q), OFun, OMsgTxt, OSess, D
'OTimStr_Dte = StrDte(D)
End Sub
