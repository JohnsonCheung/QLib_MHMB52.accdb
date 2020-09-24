Attribute VB_Name = "MxXlsLo_Loy"
Option Compare Text
Const CMod$ = "MxXlsLo_Loy."
Option Explicit

Function Loy(B As Workbook) As ListObject() ' return all Lo
Dim S As Worksheet: For Each S In B.Sheets
    Dim L As ListObject: For Each L In S.ListObjects
        PushObj Loy, L
    Next
Next
End Function

Function LoyTbl(B As Workbook) As ListObject() ' return all Lo in @B with Name Lo_*
Dim I: For Each I In Itr(Loy(B))
    If HasPfx(CvLo(I).Name, "Lo_") Then PushObj LoyTbl, I
Next
End Function
