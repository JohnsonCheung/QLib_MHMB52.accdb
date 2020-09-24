Attribute VB_Name = "MxDao_Db_Dta"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Db_Dta."
Function DcLngQ(D As Database, Q) As Long(): DcLngQ = DcLngRs(Rs(D, Q)): End Function

Private Sub B_Rs()
Shell "Subst N: c:\subst\users\user\desktop", vbHide
Const S$ = "SELECT qSku.*" & _
" FROM [N:\SAPAccessReports\DutyPrepay5\DutyPrepay5 (With Import).accdb].[qSku] AS qSku;"
BrwAy CsyRs(Rs(DbTmp, S))
End Sub

Function DcRs(R As Dao.Recordset, Optional F = 0) As Variant():        DcRs = intoDcRs(AvEmp, R, F):                  End Function
Function DcStrRs(R As Dao.Recordset, Optional F = 0) As String():   DcStrRs = intoDcRs(SyEmp, R, F):                  End Function
Function DcIntQ(D As Database, Q) As Integer():                      DcIntQ = intoDcRs(IntyEmp, Rs(D, Q)):            End Function
Function DcStrF(D As Database, T, F$, Optional Bepr$) As String():   DcStrF = DcStrRs(Rs(D, SqlSelFf(T, F, , Bepr))): End Function
Function DcStrTFC(TF$, Optional Bepr$) As String():                DcStrTFC = DcStrTF(CDb, TF, Bepr):                 End Function
Function DcStrTF(D As Database, TF$, Optional Bepr$) As String():   DcStrTF = DcStrRs(RsTF(D, TF, Bepr)):             End Function
Private Function intoDcF(Into, D As Database, T, F$):               intoDcF = intoDcRs(Into, RsF(D, T, F)):           End Function
Private Function intoDcRs(IntoyAy, R As Dao.Recordset, Optional F = 0)
Dim O: O = IntoyAy: Erase O
With R
    If Not .EOF Then .MoveFirst
    While Not .EOF
        PushI O, .Fields(F).Value
        .MoveNext
    Wend
End With
intoDcRs = O
End Function
Function FnyQn(D As Database, Qn) As String()
On Error GoTo X
FnyQn = Itn(RsQn(D, Qn).Fields)
Exit Function
X:
Dim E$: E = Err.Description
Inf CSub, "Error in open Qn", "Er Db Qn SqlQn", E, D.Name, Qn, SqlQn(D, Qn)
End Function
Function FnyQC(Q) As String(): FnyQC = FnyQ(CDb, Q): End Function
Function FnyQ(D As Database, Q) As String()
On Error GoTo X
FnyQ = Itn(Rs(D, Q).Fields)
Exit Function
X:
Dim E$: E = Err.Description
Inf CSub, "Error in open Qn", "Er Db Q", E, D.Name, Q
End Function
Function FnyQnC(Q) As String(): FnyQnC = FnyQn(CDb, Q): End Function
Private Sub B_FnyQ()
Const S$ = "SELECT * from Sku"
DmpAy FnyQC(S)
End Sub

Private Sub B_DrQ():                               MsgBox JnSpc(DrQ(CDb, "Select YY,MM,DD from [OH] where YY=20")): End Sub
Function DrQ(D As Database, Q) As Variant(): DrQ = DrRs(Rs(D, Q)):                                                  End Function
Function DyQ(D As Database, Q) As Variant(): DyQ = DyRs(Rs(D, Q)):                                                  End Function
Function DyRs(R As Dao.Recordset, Optional IsIncFldn As Boolean) As Variant()
If IsIncFldn Then
    PushI DyRs, FnyRs(R)
End If
If Not HasRec(R) Then Exit Function
With R
    .MoveFirst
    While Not .EOF
        PushI DyRs, DrRs(R)
        .MoveNext
    Wend
    .MoveFirst
End With
End Function
Function DcStrQ(D As Database, Q) As String():  DcStrQ = DcStrRs(Rs(D, Q)): End Function
Function DcStrQC(Q) As String():               DcStrQC = DcStrQ(CDb, Q):    End Function
