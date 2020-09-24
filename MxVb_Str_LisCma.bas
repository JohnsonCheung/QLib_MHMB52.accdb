Attribute VB_Name = "MxVb_Str_LisCma"
Option Compare Text
Const CMod$ = "MxVb_Str_LisCma."
Option Explicit

Function SyCmaLis(CmaLis$) As String()
Const CSub$ = CMod & "SyCmaLis"
Dim Posy%()
    Dim Lvl%
    Dim J%: For J = 1 To Len(CmaLis)
        Select Case Mid(CmaLis, J, 1)
        Case ",": If Lvl = 0 Then PushI Posy, J
        Case "(": Lvl = Lvl + 1
        Case ")": Lvl = Lvl - 1: If Lvl < 0 Then Thw CSub, "Invalid CmaLis (Lvl is -ve)", "CmaLis Lvl", CmaLis, Lvl
        Case Else
        End Select
    Next
    If Lvl <> 0 Then
        Thw CSub, "Invalid CmaLis (Lvl <> 0 at end)", "CmaLis Lvl", CmaLis, Lvl
    End If
SyCmaLis = SyPosy(CmaLis, Posy)
End Function
