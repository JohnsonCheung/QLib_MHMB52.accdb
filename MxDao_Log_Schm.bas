Attribute VB_Name = "MxDao_Log_Schm"
Option Compare Text
Const CMod$ = "MxDao_Log_Schm."
Option Explicit
Private X_DbLog As Database

Sub ClsLog(): ClsDb X_DbLog: End Sub
Sub DltFbLog()
ClsLog
If HasFfn(FbLog) Then
    Kill FbLog
    Debug.Print "DltFbLog: FbLog-[" & FbLog & "] is deleted"
Else
    Debug.Print "DltFbLog: FbLog-[" & FbLog & "] not exist"
End If
End Sub

Function DbLog() As Database
If Not IsDbOk(X_DbLog) Then
    If Not HasFfn(FbLog) Then
        CrtFb FbLog
        W3CrtTbLog4
    End If
    Set X_DbLog = Db(FbLog)
End If
Set DbLog = X_DbLog
End Function
Private Sub W3CrtTbLog4()
Dim F$: F = FbLog
DltFfnIf F
CrtFb F
Dim D As Database: Set D = Db(F)
If False Then
    W3CrtTbLog4Schm D
Else
    W3CrtTbLog4Nrm D
End If
End Sub
Private Sub W3CrtTbLog4Schm(D As Database)
SchmCrt DbLog, W3Schm
End Sub
Private Function W3Schm() As String()
Erase XX
X "E Mem | Mem Req AlZZLen"
X "E Txt | Txt Req"
X "E Crt | Dte Req Dft=Now"
X "E Dte | Dte"
X "E Amt | Cur"
X "F Amt * | *Amt"
X "F Crt * | CrtDte"
X "F Dte * | *Dte"
X "F Txt * | Fun * Txt"
X "F Mem * | Lines"
X "T Sess | * CrtDte"
X "T Msg  | * Fun *Txt | CrtDte"
X "T Log   | * Sess Msg CrtDte"
X "T LogV  | * Log Lines"
X "D . Fun | Function name that call the log"
X "D . Fun | Function name that call the log"
X "D . Msg | it will a new record when Log-function is first time using the Fun+MsgTxt"
X "D . Msg | ..."
W3Schm = XX
Erase XX
End Function
Private Sub W3CrtTbLog4Nrm(D As Database)
Dim T As Dao.TableDef
Set T = New Dao.TableDef
T.Name = "LogSess"
AddFdId T
AddFdTimstmp T, "Dte"
D.TableDefs.Append T
'
Set T = New Dao.TableDef
T.Name = "LogMsg"
AddFdId T
AddFdLng T, "FunId"
AddFdTxt T, "Msg"
AddFdTimstmp T, "Dte"
T.Indexes.Append IdxNwSk(T, Sy("FunId", "Msg"))
T.Indexes.Append IdxNwPk(T)
D.TableDefs.Append T
'
Set T = New Dao.TableDef
T.Name = "LogFun"
AddFdId T
AddFdTxt T, "Fun"
AddFdTimstmp T, "Dte"
T.Indexes.Append IdxNwPk(T)
T.Indexes.Append IdxNwSk(T, Sy("Fun"))
D.TableDefs.Append T
'
Set T = New Dao.TableDef
T.Name = "Log"
AddFdId T
AddFdLng T, "LogSessId"
AddFdLng T, "LogMsgId"
AddFdLng T, "LogFunId"
AddFdTimstmp T, "Dte"
D.TableDefs.Append T
'
Set T = New Dao.TableDef
T.Name = "LogLines"
AddFdId T
AddFdLng T, "LogId"
AddFdMem T, "Lines"
D.TableDefs.Append T
End Sub

Function FbLog$():        FbLog = PthLog & X_Fn: End Function
Private Function X_Fn$():  X_Fn = "Log.accdb":   End Function
Function PthLog$()
Static Y$: If Y = "" Then Y = PthAddFdrEns(PthAssPC, ".Log")
PthLog = Y
End Function
