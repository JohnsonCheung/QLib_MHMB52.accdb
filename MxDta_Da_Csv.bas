Attribute VB_Name = "MxDta_Da_Csv"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Csv."

Function DrCsl(Csl) As Variant()
Const CSub$ = CMod & "DrCsl"
If Not HasQuoDbl(Csl) Then
    DrCsl = Split(Csl, ",")
    Exit Function
End If
Dim L$: L = Trim(Csl)
Dim J%
While L <> ""
    ThwLoopTooMuch CSub, J
    PushI DrCsl, TmCsvShf(L)
Wend
End Function

Function SyCsl(Csl) As String()
Const CSub$ = CMod & "SyCsl"
If Not HasQuoDbl(Csl) Then
    SyCsl = SplitCma(Csl)
    Exit Function
End If
Dim L$: L = Trim(Csl)
Dim J%
While L <> ""
    ThwLoopTooMuch CSub, J
    PushI SyCsl, TmCsvShf(L)
Wend
End Function

Function TmCsvShf$(OLn$)
Const CSub$ = CMod & "TmCsvShf"
Dim NotDblQ As Boolean
Dim DblQCommaPos As Boolean
Dim LasIsDblQ As Boolean
    NotDblQ = ChrFst(OLn) <> vbQuoDbl
    If Not NotDblQ Then DblQCommaPos = InStr(2, OLn, vbQuoDbl & vbCma)
    If DblQCommaPos = 0 Then LasIsDblQ = ChrLas(OLn) = vbQuoDbl

Select Case True
Case NotDblQ
    TmCsvShf = BefCmaOrAll(OLn)
    OLn = AftCma(OLn)
Case DblQCommaPos > 0
    TmCsvShf = Replace(Mid(OLn, 2, DblQCommaPos - 1), vbQuoDbl2, vbQuoDbl)
    OLn = Mid(OLn, DblQCommaPos + 1)
Case LasIsDblQ
    TmCsvShf = Replace(Mid(OLn, 2, DblQCommaPos - 1), vbQuoDbl2, vbQuoDbl)
    OLn = ""
Case Else
    Thw CSub, "CsvEr: OLn has ChrFst is DblQ, No DblQComm, Las<>DblQ, it should be Closing-vbQuoDbl", "OLn", OLn
End Select
End Function

Function QuoCsv$(V)
Select Case True
Case IsStr(V): QuoCsv = QuoDbl(Replace(V, vbQuoDbl, vbQuoDbl2))
Case IsDte(V): QuoCsv = "#" & Format(V, "YYYY-MM-DD HH:MM:SS") & "#"
Case IsEmpty(V):
Case Else: QuoCsv = Nz(V, vbQuoDbl2)
End Select
End Function

Function Csl$(Dr)
Dim U%: U = UB(Dr)
If U = -1 Then Exit Function
Dim J&
Dim O$(): ReDim O(U)
Dim V: For Each V In Dr
    O(J) = QuoCsv(V)
    J = J + 1
Next
Csl = JnCma(O)
End Function
