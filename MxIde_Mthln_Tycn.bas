Attribute VB_Name = "MxIde_Mthln_Tycn"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthln_Tycn."

Function TycTycn$(Tycn)
Select Case Tycn
Case "Boolean":   TycTycn = "*"
Case "String":   TycTycn = "$"
Case "Integer":  TycTycn = "%"
Case "Long":     TycTycn = "&"
Case "Double":   TycTycn = "#"
Case "Single":   TycTycn = "!"
Case "Currency": TycTycn = "@"
End Select
End Function

Function Tycn$(Tyc$)
Const CSub$ = CMod & "Tycn"
Dim O$
Select Case Tyc
Case "": O = "Variant"
Case "#": O = "Double"
Case "%": O = "Integer"
Case "!": O = "Signle"
Case "@": O = "Currency"
Case "^": O = "LongLong"
Case "$": O = "String"
Case "&": O = "Long"
Case Else: Thw CSub, "Invalid Tyc", "Tyc VdtTycLis", Tyc, LisTyc
End Select
Tycn = O
End Function
