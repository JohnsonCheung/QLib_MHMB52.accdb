Attribute VB_Name = "MxDao_Log"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Log."
Dim LogSessId&
Sub Log(Fun$, MsgTxt$, ParamArray Ap())
Dim D As Database: Set D = DbLog
If LogSessId = 0 Then LogSessId = WLogSessId(D)
Dim LogFunId&: LogFunId = WLogFunId(D, Fun)
Dim LogMsgId&: LogMsgId = WLogMsgId&(D, LogFunId, MsgTxt)
Dim LogId&: LogId = WLogId(D, LogSessId, LogMsgId, LogFunId)
Dim Av(): Av = Ap
If Si(Av) = 0 Then Exit Sub
With D.TableDefs("LogLines").OpenRecordset
    Dim V: For Each V In Av
        .AddNew
        !LogId = LogId
        !Lines = LinesV(V)
        .Update
    Next
    .Close
End With
End Sub
Private Function WLogMsgId&(D As Database, FunId&, Msg$)
With D.TableDefs("LogMsg").OpenRecordset
    .Index = "LogMsg"
    .Seek "=", FunId, Msg
    If .NoMatch Then
        .AddNew
        !FunId = FunId
        !Msg = Msg
        WLogMsgId = !LogMsgId
        .Update
    Else
        WLogMsgId = !LogMsgId
    End If
End With
End Function
Private Function WLogFunId&(D As Database, Fun$)
With D.TableDefs("LogFun").OpenRecordset
    .Index = "LogFun"
    .Seek "=", Fun
    If .NoMatch Then
        .AddNew
        !Fun = Fun
        WLogFunId = !LogFunId
        .Update
    Else
        WLogFunId = !LogFunId
    End If
End With
End Function
Private Function WLogId&(D As Database, LogSessId&, LogMsgId&, LogFunId&)
With D.TableDefs("Log").OpenRecordset
    .AddNew
    !LogSessId = LogSessId
    !LogMsgId = LogMsgId
    !LogFunId = LogFunId
    WLogId = !LogId
    .Update
End With
End Function
Private Function WLogSessId&(D As Database)
If LogSessId = 0 Then
    With D.TableDefs("LogSess").OpenRecordset
        .AddNew
        WLogSessId = !LogSessId
        .Update
        .Close
    End With
End If
End Function
Sub EndLogSess(): Log ".", "LogSessEnd": LogSessId = 0: End Sub
