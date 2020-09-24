Attribute VB_Name = "MxVb_Str_Ssub_ChrPfxAll"
Option Compare Text
Option Explicit
Private Sub B_ShfChrPfxAll()
GoSub T1
Exit Sub
Dim OLn$, ChrPfx$, C As eCas, EptOLn$
T1:
    OLn = "((((( AA"
    ChrPfx = "("
    C = eCasIgn
    Ept = "((((("
    EptOLn = " AA"
    GoTo Tst
Tst:
    Act = ShfChrPfxAll(OLn, ChrPfx, C)
    Debug.Assert Act = Ept
    Debug.Assert EptOLn = OLn
    Return
End Sub
Function RmvChrPfxAll$(S, ChrPfx$, Optional C As eCas): RmvChrPfxAll = Mid(S, WLenChrPfxAll(S, ChrPfx, C) + 1): End Function
Function ShfChrPfxAll$(OLn$, ChrPfx$, Optional C As eCas)
Dim L%: L = WLenChrPfxAll(OLn, ChrPfx, C)
ShfChrPfxAll = ShfLeft(OLn, L)
End Function
Function ShfLeft$(OLn$, NLeft%)
ShfLeft = Left(OLn, NLeft)
OLn = Mid(OLn, NLeft + 1)
End Function
Private Function WLenChrPfxAll%(S, ChrPfx$, Optional C As eCas)
Dim J%: For J = 1 To Len(S)
    If Not IsEqStr(Mid(S, J, 1), ChrPfx, C) Then WLenChrPfxAll = J - 1: Exit Function
Next
WLenChrPfxAll = Len(S)
End Function
