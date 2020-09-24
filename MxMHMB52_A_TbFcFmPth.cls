VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_TbFcFmPth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'1-Good-MH-FcFxFn return 1-StmYm
'1-Good-UD-FcFxFn return 2-StmYm
'Fc 20??-?? MH 8600.xlsx
'Fc 20??-?? MH 8700.xlsx
'Fc 20??-?? UD.xlsx
'         1
'123456789012345678
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_FcTb_FmPth."
Private MHNw As New MxMHMB52_MHNw
Private Fun As New MxMHMB52_MHFun
Private Sub B_WStmYmyzPth()
Dim A()  As StmYm: A = WStmYmyzPth(MH.FcIO.Pthi)
Stop
End Sub
Sub RfhFmPthi(): RfhFmPth MH.FcIO.Pthi: End Sub
Sub RfhFmPth(Pth$)
'Aim: Create new record to table-Fc according the Fc-Import-Pth$
Dim A() As StmYm: A = WStmYmyzPth(Pth)
Dim M As StmYm
Dim J%: For J = 0 To UbStmYm(A)
    M = A(J)
    Dim Sql$: Sql = "Select * from Fc" & MH.FcWh.WhStmYm(M)
    With RsQC(Sql)
        If .EOF Then
            .AddNew
            !VerYY = M.Y
            !VerMM = M.M
            Dim Fxi$: Fxi = MH.FcIO.Fxi(M)
            !Siz = FileLen(Fxi)
            !Tim = FileDateTime(Fxi)
            !DteLoad = Null
            !Stm = M.Stm
            .Update
        End If
    End With
Next
If IsFrmnOpn("LoadFc") Then Form_LoadFc.Requery
End Sub
Private Function WStmYmyzPth(Pth$) As StmYm()
Dim FfnFc: For Each FfnFc In Itr(MH.FcIO.FxyInp)
    PushStmYm WStmYmyzPth, WStmYmzFnFc(Fn(FfnFc))
Next
End Function
Private Function WStmYmzFnFc(FnFc) As StmYm
Dim Y As Byte, M As Byte, Stm$
    Stm = WStmFnFc(FnFc): If Stm = "" Then Exit Function
    Y = WYyFnFc(FnFc): If Y = 0 Then Exit Function
    M = WMmFnFc(FnFc): If M = 0 Then Exit Function
WStmYmzFnFc = MHNw.StmYm(Stm, Y, M)
End Function
Private Function WYyFnFc(FnFc) As Byte
On Error Resume Next
WYyFnFc = Mid(FnFc, 6, 2)
End Function
Private Function WMmFnFc(FnFc) As Byte
Dim A%: A = Val(Mid(FnFc, 9, 2))
If 1 <= A And A <= 12 Then WMmFnFc = A
End Function
Private Function WStmFnFc$(FnFc): WStmFnFc = Fun.StmStm2(Mid(FnFc, 12, 2)): End Function
