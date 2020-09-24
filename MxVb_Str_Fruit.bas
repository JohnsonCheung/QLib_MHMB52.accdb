Attribute VB_Name = "MxVb_Str_Fruit"
'#Qtp:Question-Mark-Template# It is a string which will be used as a template
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Fruit."
Function SyExpandSS(Qtp$, SeedSS$) As String(): Stop 'SyExpandSS = SyExpand(Qtp, Tml(SeedSS)): End Function
End Function
Function SyExpand(Qtp$, Seedy$()) As String()
Dim S: For Each S In Itr(Seedy)
    PushI SyExpand, RplQ(Qtp, S)
Next
End Function

Function LsyExpandSS(Qvbl$, SeedSS$) As String(): Stop 'LsyExpandSS = LsyExpand(Qvbl, Tml(SeedSS)): End Function

End Function
Function LsyExpand(Qvbl$, Seedy$()) As String()
Dim T$: T = RplVBar(Qvbl)
Dim S: For Each S In Itr(Seedy)
    PushI LsyExpand, RplQ(T, S)
Next
End Function

Private Sub B_LsyExpandSS()
GoSub ZZ
Exit Sub
Dim Qvbl$, Seedy$()
ZZ:
    Erase XX
    X "Sub Push?(O() As ?, M As ?)"
    X "Dim N&"
    X "N = ?Si(O)"
    X "ReDim Preserve O(N)"
    X "O(N) = M"
    X "End Sub"
    X ""
    X "Function ?SI&(A() As ?)"
    X "On Error Resume Next"
    X "?Si = Ubound(A) + 1"
    X "End Function"
    X ""
    X ""
    Qvbl = JnVBar(XX)
    Erase XX
    Brw LsyExpandSS(Qvbl, "S12 XX")
T0:
    Qvbl = "Sub Tst?()|Dim A As New ?: A.Tst|End Sub"
    Seedy = SySs("Xws Xwb Xfx Xrg")
    Erase XX
    X ""
    X ""
    Ept = JnCrLf(XX)
    Erase XX
    GoTo Tst
Tst:
    Act = LsyExpand(Qvbl, Seedy)
    C
    Return
End Sub
