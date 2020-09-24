Attribute VB_Name = "MxDao_Att_AttTb2_ChkAttTb2"
Option Compare Text
Const CMod$ = "MxDao_Att_AttTb2_AttChkTb2."
Option Explicit

Sub ChkAttTb2C(): ChkAttTb2 CDb: End Sub
Sub ChkAttTb2(D As Database)
Const CSub$ = CMod & "ChkAttTb2"
Dim E$()
    Dim E1$(): E1 = WErMisAttf_InTbAtt(D)
    Dim E2$(): E2 = WErMisAttf_InTbAttd(D)
    Dim E3$(): E3 = MsgyTbNoParOrChd(D, "Att", "Attd", "AttId")
    E = SyAddAp(E1, E2, E3)
If Si(E) = 0 Then Exit Sub
Dim Tit$
    PushI Tit, "Database tables [Att] & [Attd] has errors"
    PushI Tit, "Database file: [" & D.Name & "]"

ChkEry E, CSub, Tit
End Sub
Private Function WErMisAttf_InTbAtt(D As Database) As String()
Const C$ = "Att.Att miss Attachment AttId[?] Attn[?] Attf[?]"
Dim DotnyAtt$()
Dim DotnyAttd$()
End Function
Private Function WErMisAttf_InTbAttd(D As Database) As String()

End Function
Function MsgyTbNoParOrChd(D As Database, TbPar, TbChd, Fldn$) As String()

End Function
