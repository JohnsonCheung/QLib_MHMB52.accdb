VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_TbReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_RfhTbReportFmPth."

Sub RfhFmPthi(): RfhFmPth MH.MB52IO.Pthi: End Sub
Sub RfhFmPth(P$)

'Aim: Detect any new MHMB52Io.Fxi in Pm.MB52Pthi folder and add record to table-Report (Y M D ..)
'     When add record to Report: Fields with value is: Y M D.
'     Assume DteCrt has default value Now()
'     Refresh Form_Rpt.Requery
Dim Ymd() As Ymd: Ymd = WYmdyPth(P)
Dim N%
Dim J%, M As Ymd: For J = 0 To UbYmd(Ymd)
    If WIns(Ymd(J)) Then
        N = N + 1
    End If
Next
If N > 0 Then
    Sts "[" & N & "] is inserted in Tbl-Report"
End If
End Sub

Private Function WYmdyPth(Pth$) As Ymd()
Const c_FSPec_MB52$ = "MB52 20??-??-??.XLSX"
Dim Fn$(): Fn = Fnay(Pth, c_FSPec_MB52)
Dim I: For Each I In Itr(Fn)
    PushYmd WYmdyPth, WYmd(I)
Next
End Function
Private Function WIns(A As Ymd) As Boolean
'Ins to Tbl-Report if @A is new
With CurrentDb.OpenRecordset("Select YY,MM,DD from Report" & MH.TbOH.WhYmd(A))
    If .EOF Then
        .AddNew
        !YY = A.Y
        !MM = A.M
        !DD = A.D
        .Update
        Sts "Inserting Tbl-Report: " & HypymdYmd(A)
        WIns = True
    End If
End With
End Function
Private Function WYmd(FnMB52) As Ymd
If WIsMB52Fn(FnMB52) Then
    With WYmd
        .Y = Mid(FnMB52, 8, 2)
        .M = Mid(FnMB52, 11, 2)
        .D = Mid(FnMB52, 14, 2)
    End With
End If
End Function
Private Function WIsMB52Fn(Fn) As Boolean
Select Case True
Case Not WHasMBPfx(Fn), Not WHasMBDte(Fn), Not IsXlsx(Fn): Exit Function
End Select
WIsMB52Fn = True
End Function
Private Function WHasMBPfx(Fn) As Boolean: WHasMBPfx = HasPfx(Fn, "MB52 20"):  End Function
Private Function WHasMBDte(Fn) As Boolean: WHasMBDte = IsDate(Mid(Fn, 6, 10)): End Function
