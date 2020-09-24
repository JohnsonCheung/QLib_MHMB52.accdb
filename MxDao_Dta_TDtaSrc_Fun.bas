Attribute VB_Name = "MxDao_Dta_TDtaSrc_Fun"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Dta_TDtaSrc_Fun."

Function TnyTDtaSrc(S As TDtaSrc) As String(): TnyTDtaSrc = TnyTF(S.TF): End Function
Function TnyTF(T() As TF) As String()
Dim J%: For J = 0 To UbTF(T)
    PushI TnyTF, T(J).Tbn
Next
End Function

Function TFwTny(T() As TF, Tny$()) As TF()
Dim J%: For J = 0 To UbTF(T)
    If HasEle(Tny, T(J).Tbn) Then
        PushTF TFwTny, T(J)
    End If
Next
End Function
