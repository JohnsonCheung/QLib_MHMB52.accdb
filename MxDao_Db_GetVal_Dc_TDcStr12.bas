Attribute VB_Name = "MxDao_Db_GetVal_Dc_TDcStr12"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Db_GetVal_Dc_TDcStr12."
Type TDcStr12: Dc1() As String: Dc2() As String: End Type
Sub AsgTDcStr12(C As TDcStr12, ODc1$(), ODc2$()): ODc1 = C.Dc1: ODc2 = C.Dc2: End Sub

Function TDcStr12T(D As Database, T, F12$, Optional IsDis As Boolean, Optional Bepr$) As TDcStr12:  TDcStr12T = TDcStr12TQ(D, SqlSelFf(T, F12, IsDis, Bepr)): End Function
Function TDcStr12TQ(D As Database, Q) As TDcStr12:                                                 TDcStr12TQ = TDcStr12Rs(Rs(D, Q)):                         End Function
Function TDcStr12Rs(R As Dao.Recordset) As TDcStr12
Dim O As TDcStr12
With R
    If Not .EOF Then .MoveFirst
    While Not .EOF
        PushI O.Dc1, Nz(.Fields(0).Value, "")
        PushI O.Dc2, Nz(.Fields(1).Value, "")
        .MoveNext
    Wend
End With
TDcStr12Rs = O
End Function
