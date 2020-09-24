Attribute VB_Name = "MxVb_Dta_Sq_Str"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Sq_Str."

Function LySq(Sq()) As String()
':SqStr: :S  ! it is Lines of StrVal Ln.
'            ! #StrVal-Ln is :StrVal separaterd by vbTab with ending vbTab.
'            ! the fst StrVal-Ln will not trim the ending vbTab, because it is used to determine how many col.
'            ! if any non-Lin1-StrVal-Ln has more fld than lin1-StrVal-Ln-fld, the extra fld are ignored and inf (this is done in %SqS%
'            ! the reverse fun is %SqStr @@
Dim IR&: For IR = 1 To UBound(Sq, 1)
    PushI LySq, JnTab(DrSq(Sq, IR))
Next
End Function
Function SqStr(SqStr_$) As Variant()
'Ret : :Sq from :SqStr
Dim Ry$(): Ry = SplitCrLf(SqStr_): If Si(Ry) = 0 Then Exit Function
Dim NR&: NR = Si(Ry)
Dim R1$: R1 = Ry(0)
Dim NC%: NC = Si(SplitTab(R1))
Dim O(): ReDim O(1 To NR, 1 To NC)
Dim IR&, IC%
Dim R: For Each R In Ry
    IR = IR + 1
    Dim C: For Each C In SplitTab(R)
        IC = IC + 1
        If IC > NC Then Exit For ' ign the extra fld, if it has more fld then lin1-fld-cnt
        O(IR, IC) = ValChrStr(C)
    Next
Next
End Function

Function ValChrStr(ChrStr)
'Ret : ! a val (Str|Dbl|Bool|Dte|Empty) fm @ChrStr.
'      ! If fst letter is
'      !   ['] is a str wi \r\n\t
'      !   [D] is a str of date, if cannot convert to date, ret empty and debug.print msg.
'      !   [T] is true
'      !   [F] is false
'      !   rest is dbl, if cannot convert to dbl, ret empty and debug.print msg @@
Dim F$: F = ChrFst(ChrStr)
Dim O$
Select Case F
Case "'": O = UnslashCrLfTab(RmvFst(ChrStr))
Case "T": O = True
Case "F": O = False
Case "D": O = CvDte(RmvFst(ChrStr))
Case ""
Case Else: O = CvDbl(RmvFst(ChrStr))
End Select
ValChrStr = O
End Function

Function ChrStr$(V, Optional Fun$)
Const CSub$ = CMod & "ChrStr"
':ChrStr: :S #Letter-Str# ! A str wi fst letter can-determine the str can converted to what value.
Dim T$: T = TypeName(V)
Dim O$
Select Case T
Case "String": O = "'" & SlashCrLfTab(V)
Case "Boolean": O = IIf(V, "T", "F")
Case "Integer", "Single", "Double", "Currency", "Long": O = V
Case "Date": O = "D" & V
Case Else: If Fun <> "" Then Inf CSub, "Val-of-TypeName[" & T & "] cannot cv to :ChrStr"
End Select
ChrStr = O
End Function
