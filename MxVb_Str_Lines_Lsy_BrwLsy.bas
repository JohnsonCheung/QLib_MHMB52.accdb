Attribute VB_Name = "MxVb_Str_Lines_Lsy_BrwLsy"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Lines_Lsy_BrwLsy."
Private Sub B_FmtLsy()
GoSub Z1
Exit Sub
Z1:
    BrwLsy SqyQryC
    Return
End Sub

Sub VcLsy(Lsy$()):  VcAy FmtLsy(Lsy):  End Sub
Sub BrwLsy(Lsy$()): BrwAy FmtLsy(Lsy): End Sub
Function FmtLsy(Lsy$()) As String()
If Si(Lsy) = 0 Then Exit Function
Dim W0%, W1%, Sep$
    W0 = NDig(Si(Lsy))
    W1 = WdtLsy(Lsy)
Dim O$(): ReDim O(UB(Lsy))
Dim Ix&: For Ix = 0 To UB(Lsy)
    O(Ix) = WFmtlLines(Lsy(Ix), W0, W1, Ix + 1) '<==2
Next
FmtLsy = WLyAddSepln(O, W0, W1)
End Function
Private Function WFmtlLines$(Lines, W0%, W1%, Ix&) ' Fmt one @Lines with @W1 optionally having @Ix depends on @W0
Dim Ly$()
If Lines = "" Then
    PushI Ly, ""
Else
    Ly = SplitCrLf(Lines)
End If
Dim O$()
PushI O, "| " & AliR(Ix, W0) & " | " & AliL(Ly(0), W1) & " |"    '<==1
Dim S$: S = "| " & Space(W0) & " | "
Dim J%: For J = 1 To UB(Ly)
    PushI O, S & AliL(Ly(J), W1) & " |" '<==2
Next
WFmtlLines = JnCrLf(O)
End Function
Private Function WIsBrkyAft(Lsy$()) As Boolean()
'@@Ret Is Brk aft ele of @Lsy.
'      Will be true is either NLn1 or NLn2 > 1  !WIsBrk
'      Same Si as @Lsy
'      Las ele is always true
Dim U&: U = UB(Lsy)
Dim O() As Boolean: ReDim O(U)
O(U) = True     'LasEle is always true
Dim IsSnglny() As Boolean: IsSnglny = WIsSnglny(Lsy)
Dim J&: For J = 0 To U - 1
    If WIsBrk(IsSnglny(J), IsSnglny(J + 1)) Then
        O(J) = True
    End If
Next
WIsBrkyAft = O
End Function
Private Function WIsBrk(IsSngln1 As Boolean, IsSngln2 As Boolean) As Boolean
WIsBrk = Not IsSngln1 Or Not IsSngln2
End Function
Private Function WIsSnglny(Lsy$()) As Boolean()
Dim Lines: For Each Lines In Lsy
    PushI WIsSnglny, NLn(Lines) <= 1
Next
End Function
Private Function WLyAddSepln(Lsy$(), W0%, W1%) As String()
Dim Sepln$: Sepln = SeplnWdty(Inty(W0, W1), eTblFmtTb)
Dim IsBrkyAft() As Boolean: IsBrkyAft = WIsBrkyAft(Lsy)
Dim OLy$()
PushI OLy, Sepln
Dim U&: U = UB(Lsy)
Dim J&: For J = 0 To U
    PushI OLy, Lsy(J)
    If IsBrkyAft(J) Then PushI OLy, Sepln
Next
WLyAddSepln = OLy
End Function

Function WdtLsy%(Lsy$())
Dim O%
Dim Lines: For Each Lines In Itr(Lsy)
    O = Max(O, WdtLines(Lines))
Next
WdtLsy = O
End Function
