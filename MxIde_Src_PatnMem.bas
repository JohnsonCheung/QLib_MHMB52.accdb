Attribute VB_Name = "MxIde_Src_PatnMem"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_PatnMem."
Enum eHsh: eHshInl: eHshExl: End Enum
Private Sub B_Memny()
Dim A$(), B$(), S$
S = SrclPC
A = Memny(S)
B = Memny(S, eHshExl)
If Si(A) <> Si(B) Then Stop
Dim O$()
Dim J%: For J = 0 To UB(A)
    PushI O, A(J) & " " & B(J)
Next
Brw FmtT1ry(O)
End Sub
Function HasMemn(S) As Boolean: HasMemn = HasRx(S, WRx):  End Function
Function Memn$(S):                 Memn = SsubRx(S, WRx): End Function '#memonic-name# a name between 2 hashChr
Private Function WRx() As RegExp
Static X As RegExp: If IsNothing(X) Then Set X = Rx("/#([A-Za-z][\w\.-]*)#/g")
Set WRx = X
End Function
Function Memny(S$, Optional H As eHsh) As String()
If H = eHshExl Then
    Memny = SsubyRx(S, WRx, 1)
Else
    Memny = SsubyRx(S, WRx)
End If
End Function
