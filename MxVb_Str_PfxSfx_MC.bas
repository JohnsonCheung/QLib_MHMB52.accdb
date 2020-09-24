Attribute VB_Name = "MxVb_Str_PfxSfx_MC"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_PfxSfx_MC."
Function StrPfx$(S, Pfx):            StrPfx = Pfx & S:       End Function
Function StrPfxSfx$(S, Pfx, Sfx): StrPfxSfx = Pfx & S & Sfx: End Function
Function StrSfx$(S, Sfx):            StrSfx = S & Sfx:       End Function

Function Sfx$(S, Suffix$, Optional C As VbCompareMethod = vbBinaryCompare)
If HasSfx(S, Suffix, C) Then Sfx = Suffix
End Function

Function IsAllEleHasPfx(Ay, Pfx, Optional C As eCas) As Boolean
If Si(Ay) = 0 Then Exit Function
Dim V: For Each V In Itr(Ay)
    If Not HasPfx(V, Pfx, C) Then Exit Function
Next
IsAllEleHasPfx = True
End Function

Function HasPfx(S, Pfx, Optional C As eCas) As Boolean:
HasPfx = StrComp(Left(S, Len(Pfx)), Pfx, VbCprMth(C)) = 0
End Function
Function HasPfxSpc(S, Pfx, Optional C As eCas) As Boolean: HasPfxSpc = HasPfx(S, Pfx & " ", C): End Function ' Does @S have pfx which is after adding ' ' to end of ele of @Pfxy
Function HasPfxSfx(S, Pfx, Sfx, Optional C As VbCompareMethod = vbTextCompare) As Boolean
If Not HasPfx(S, Pfx, C) Then Exit Function
If Not HasSfx(S, Sfx, C) Then Exit Function
HasPfxSfx = True
End Function

Function HasSfxy(S, Sfxy$(), Optional C As eCas) As Boolean
Dim Sfx: For Each Sfx In Itr(Sfxy)
    If HasSfx(S, Sfx, C) Then HasSfxy = True: Exit Function
Next
End Function
Function HasPfxy(S, Pfxy$(), Optional C As eCas) As Boolean
Dim Pfx: For Each Pfx In Itr(Pfxy)
    If HasPfx(S, Pfx, C) Then HasPfxy = True: Exit Function
Next
End Function
Private Sub B_NSpcPfx()
GoSub T1
GoSub T2
GoSub T3
Exit Sub
Dim S
T1:
    S = "    123"
    Ept = 4
    GoTo Tst
T3:
    S = " "
    Ept = 1
    GoTo Tst
T2:
    S = "13"
    Ept = 0
    GoTo Tst
Tst:
    Act = NSpcPfx(S)
    Debug.Assert Act = Ept
    Return
End Sub
Function NSpcPfx%(S)
Static R As RegExp: If IsNothing(R) Then Set R = Rx(" +")
Dim M As Match: Set M = Mch(S, R): If IsNothing(M) Then Exit Function
NSpcPfx = M.Length
End Function

Function HasSfx(S, Sfx, Optional C As eCas) As Boolean: HasSfx = IsEqStr(Right(S, Len(Sfx)), Sfx, C): End Function
Function PfxPfxy$(S, Pfxy) ' ret one of *pfxEle of @Pfxy if @S has such *pfxEle else blnk.
'@Pfxy cannot be $(), becuase, it !Pfxy will be called by !PfxAp
Dim P: For Each P In Pfxy
    If HasPfx(S, P) Then PfxPfxy = P: Exit Function
Next
End Function

Function PfxPfxySpc$(S, Pfxy, Optional C As eCas) ' ret one of *pfxEle of @Pfxy if @S has such *pfxEle+spc else blnk
'@Pfxy cannot be Sy, becuase, it !Pfxy will be called by !PfxAp
Dim P: For Each P In Pfxy
    If HasPfx(S, P & " ") Then PfxPfxySpc = P: Exit Function
Next
End Function

Function SfxSfxySpc$(S, Sfxy$()) ' ret one of *sfxEle of @Sfxy if @S has such sfx+spc or blnk
Dim Sfx: For Each Sfx In Sfxy
    If HasSfxSpc(S, Sfx) Then SfxSfxySpc = Sfx: Exit Function
Next
End Function

Function HasSfxSpc(S, Sfx, Optional C As eCas) As Boolean: HasSfxSpc = HasSfx(S, Sfx + " ", C): End Function
