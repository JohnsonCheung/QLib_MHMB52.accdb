Attribute VB_Name = "MxVb_Str_Bkt_Fun"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Bkt_Fun."

Function PosBktCls%(S, PosBktOpn%, Optional BktOpn$ = vbBktOpn)
Const CSub$ = CMod & "PosBktCls"
If Mid(S, PosBktOpn, 1) <> BktOpn Then ThwPm CSub, "@PosBktOpn not point to a @BktOpn", "Chr-At-PosBktOpn BktOpn PosBktOpn S", Mid(S, PosBktOpn, 1), BktOpn, PosBktOpn, S
Dim BktC$: BktC = BktCls(BktOpn)
Dim NOpn%
Dim J%: For J = PosBktOpn + 1 To Len(S)
    Select Case Mid(S, J, 1)
    Case BktC
        If NOpn = 0 Then
            PosBktCls = J '<===
            Exit Function
        End If
        NOpn = NOpn - 1
    Case BktOpn
        NOpn = NOpn + 1
    End Select
Next
Stop
End Function

Function BktCls$(BktOpn$)
Select Case BktOpn
Case "(": BktCls = ")"
Case "[": BktCls = "]"
Case "{": BktCls = "}"
Case Else: Stop
End Select
End Function
