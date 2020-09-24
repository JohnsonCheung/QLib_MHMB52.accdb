Attribute VB_Name = "MxIde_Mthln_MthTy_ShtMthTy"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthln_MthTy_ShtMthTy."

Function ShtMthTy$(Mtht)
Dim O$
Select Case Mtht
Case "Property Get": O = "Get"
Case "Property Set": O = "Set"
Case "Property Let": O = "Let"
Case "Function":     O = "Fun"
Case "Sub":          O = "Sub"
End Select
ShtMthTy = O
End Function

Function MthtSht$(ShtMthTy)
Const CSub$ = CMod & "MthTySht"
Dim O$
Select Case ShtMthTy
Case "Get": O = "Property Get"
Case "Set": O = "Property Set"
Case "Let": O = "Property Let"
Case "Fun": O = "Function"
Case "Sub": O = "Sub"
Case Else:  Thw CSub, "Given ShtMthTy is invalid", "ShtMthTy Invalid-ShtMthTy", ShtMthTy, "Get Set Let Fun Sub"
End Select
MthtSht = O
End Function

Function Shtmthkd$(Shtmtht$)
Dim O$
Select Case Shtmtht
Case "Get": O = "Prp"
Case "Set": O = "Prp"
Case "Let": O = "Prp"
Case "Fun": O = "Fun"
Case "Sub": O = "Sub"
Case Else: Thw CSub, "Invalid ShtMthTy", "ShtMthTy", Shtmtht
End Select
Shtmthkd = O
End Function


Private Sub B_Mthkd()
Dim A$
Ept = "Property": A = "Property Get": GoSub Tst
Ept = "Property": A = "Property Get":         GoSub Tst
Ept = "Property": A = " Property Get":        GoSub Tst
Ept = "Property": A = "Friend Property Get":  GoSub Tst
Ept = "Property": A = "Friend  Property Get": GoSub Tst
Ept = "":         A = "FriendProperty Get":   GoSub Tst
Exit Sub
Tst:
    Act = Mthkd(A)
    C
    Return
End Sub

Function Mthkd$(Mtht$)
Select Case Mtht
Case "Property Get", "Property Set", "Property Let": Mthkd = "Property"
Case "Sub": Mthkd = "Sub"
Case "Function": Mthkd = "Function"
End Select
End Function

Function MthkdL$(Ln): MthkdL = Mthkd(MthTyLn(Ln)): End Function
