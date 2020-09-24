Attribute VB_Name = "MxVb_Dta_Lndy"
Option Compare Text
Option Explicit

Function FmtLndy(Lndy(), Optional CiiAliR$, Optional CiiHypKey$, Optional CiiNbr$, Optional NoIx As Boolean, Optional Sep$ = " ", Optional NHdr%) As String()
If Si(Lndy) = 0 Then Exit Function
Dim ODy(), Dy1(), Dy2(), Dy3(), Dy4(), Dy5()
    Dy1 = DyReDimDc(Lndy)
    Dy2 = DyCvNbr(Dy1, CiiNbr)
    Dy3 = DyAli(Dy2, IntySs(CiiAliR))
    Dy4 = DyAddDcIx(Dy3, NoIx, NHdr)
    Dy5 = DySrtCii(Dy4, CiiHypKey)
FmtLndy = DyJn(Dy5, Sep:=Sep)
BrwAy FmtLndy, "FmtLndy ": Stop
End Function

Sub BrwLndy(Lndy(), Optional CiiAliR$, Optional CiiHypKey$, Optional CiiNbr$, Optional NoIx As Boolean, Optional Sep$ = " ")
BrwAy FmtLndy(Lndy, CiiAliR, CiiHypKey, CiiNbr, NoIx, Sep)
End Sub
Private Function DyCvNbr(Dy(), CiiNbr$)
If CiiNbr = "" Then DyCvNbr = Dy: Exit Function
Dim Ciy%()
    Dim I: For Each I In SplitSpc(CiiNbr)
        PushI Ciy, I
    Next
Dim O()
    ReDim O(UB(Dy))
    Dim Dr, J&: For Each Dr In Dy
        Dim Cix: For Each Cix In Ciy
            Dr(Cix) = Val(Dr(Cix))
        Next
        O(J) = Dr
        J = J + 1
    Next
DyCvNbr = O
End Function
