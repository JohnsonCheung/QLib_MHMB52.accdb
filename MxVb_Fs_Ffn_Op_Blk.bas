Attribute VB_Name = "MxVb_Fs_Ffn_Op_Blk"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_Ffn_Op_Blk."
Const C_LenBlk% = 8192
Private Sub B_BlkFfn()
Dim T$, S$, A$
S = "sllksdfj lsdkjf skldfj skldfj lk;asjdf lksjdf lsdkfjsdkflj "
T = FtTmp
WrtStr S, T
Debug.Assert FfnLen(T) = Len(S)
A = BlkFfn(T, 1)
Debug.Assert A = Left(S, 128)
End Sub
Function BlkFfn$(Ffn, IBlk)
Dim F%: F = FnoRnd(Ffn, 128)
BlkFfn = BlkFno(F, IBlk)
Close #F
End Function
Function BlkFstN$(Fno%, N&)
If N <= 0 Then Exit Function
Dim O$(): ReDim O(N - 1)
Dim A As String * C_LenBlk, J&: For J = 1 To N
    Get #Fno, J, A
    PushI O, A
Next
BlkFstN = Join(O, "")
End Function

Function BlkLas$(Fno%, NCmplBlk&, ZLenBlkLas%)
If ZLenBlkLas = 0 Then Exit Function
Dim A As String * C_LenBlk: Get #Fno, NCmplBlk + 1, A
BlkLas = Left(A, ZLenBlkLas)
End Function

Private Function ZLenBlkLas%(Si&): ZLenBlkLas = Si - ((NBlk(Si, C_LenBlk) - 1) * C_LenBlk): End Function
Function BlkFno$(Fno%, IBlk)
Dim A As String * 128
Get #Fno, IBlk, A
BlkFno = A
End Function
