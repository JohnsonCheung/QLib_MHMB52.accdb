Attribute VB_Name = "MxVb_Ay_Op_ShfAy"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_Sft."

Function IsShfEle(OAy, Ele) As Boolean
Dim At&
Dim I: For Each I In Itr(OAy)
    If I = Ele Then
        IsShfEle = True
        OAy = AeAt(OAy, At)
        Exit Function
    End If
    At = At + 1
Next
End Function
Function ShfIntBet%(OAy, A%, B%)
Dim At&
Dim I: For Each I In Itr(OAy)
    If IsNumeric(I) Then
        If IsBet(CInt(I), A, B) Then
            ShfIntBet = I
            
            OAy = AeAt(OAy, At)
            Exit Function
        End If
    End If
    At = At + 1
Next
End Function
