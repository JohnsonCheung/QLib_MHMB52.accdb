Attribute VB_Name = "MxIde_Mthln_Mdy"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthln_Mdy."

Function MdySht$(ShtMdy)
Dim O$
Select Case ShtMdy
Case "Pub": O = "Public"
Case "Prv": O = "Private"
Case "Frd": O = "Friend"
Case ""
Case Else: Stop
End Select
MdySht = O
End Function
