Attribute VB_Name = "MxVb_Lg"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Lg."

Sub LgrBrw()
BrwFt LgrFt
End Sub

Property Get LgrFilNo%()
LgrFilNo = FnoA(LgrFt)
End Property

Property Get LgrFt$()
LgrFt = LgrPth & "Log.txt"
End Property

Sub LgrLg(Msg$)
Dim F%: F = LgrFilNo
Print #F, StrNow & " " & Msg
If LgrFilNo = 0 Then Close #F
End Sub

Property Get LgrPth$()
Dim O$:
'O = WrkPth: PthEns O
O = O & "Log\": PthEns O
LgrPth = O
End Property
