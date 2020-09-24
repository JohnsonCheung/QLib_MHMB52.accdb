Attribute VB_Name = "MxIde_Dv_Enm_zIntl_MsrcDvEnm"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Deri_Enm_Msrcdevenm."
Function TEnmBlk(EnmBlk, Optional Mdn$) As TEnm
Dim Blk$(): Blk = EnmBlk
Dim Stmt$(): Stmt = StmtySrc(Blk)
End Function

Private Sub B_WMsrcoptdvenmCmp()
GoSub ZZ
Exit Sub
ZZ:
    Dim Nlyy() As Nly
    Dim C As VBComponent: For Each C In CPj.VBComponents
        PushNlyopt Nlyy, WMsrcoptDvenmCmp(C)
    Next
    BrwNlyy Nlyy
    Return
End Sub
Function MsrcyDvEnmP(P As VBProject) As Nly()
Dim C As VBComponent: For Each C In P.VBComponents
    PushNlyopt MsrcyDvEnmP, WMsrcoptDvenmCmp(C)
Next
End Function
Private Function WMsrcoptDvenmCmp(C As VBComponent) As Nlyopt
Dim Src$(): Src = SrcCmp(C)
WMsrcoptDvenmCmp = NlyoptLyopt(LyoptOldNew(Src, SrcDvenm(Src)), C.Name)
End Function
