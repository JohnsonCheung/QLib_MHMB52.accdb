Attribute VB_Name = "MxDta_Aet_Push"
Option Compare Text
Const CMod$ = "MxDta_Aet_Push."
Option Explicit

Sub PushAet(OAet As Dictionary, Aet As Dictionary): PushAetItr OAet, Aet.Keys: End Sub
Sub PushAetAy(OAet As Dictionary, Ay)
Dim I: For Each I In Itr(Ay)
    PushAetEle OAet, I
Next
End Sub

Sub PushAetEle(OAet As Dictionary, Ele)
If Not OAet.Exists(Ele) Then OAet.Add Ele, Empty
End Sub

Sub PushAetItr(OAet As Dictionary, Itr, Optional NoBlnkStr As Boolean)
Dim I
If NoBlnkStr Then
    For Each I In Itr
        If I <> "" Then
            PushAetEle OAet, I
        End If
    Next
Else
    For Each I In Itr
        PushAetEle OAet, I
    Next
End If
End Sub
