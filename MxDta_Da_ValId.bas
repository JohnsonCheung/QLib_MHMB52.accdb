Attribute VB_Name = "MxDta_Da_ValId"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_ValId."

Function DrsAdd_valid_cnt_2Dc(D As Drs, Coln$, Optional ColnPfx$) As Drs
'Fm D       : ..@Coln..  ! must have a Str-DcDrs-@Coln
'Fm Coln    : Str-DcDrs-Nm !
'Fm ColnPfx :            ! to fnd: %C1 and %C1, where %C1 = %P%Id & %C2 = %P%Cnt where %P = @ColnPfx & @Coln
'Ret        : .. %C1 %C2 ! Add 2 col: ValId-&-Cnt-col: %C1 & %C2 at end of @D.
'                        ! %C1 is ValId-DcDrs running fm 1 for each dist val of *@Coln
'                        ! %C2 is Cnt-DcDrs is the cnt of occurrance such id.  Rec of sam @Coln-Val will have sam *Cnt
Dim P$:   P = ColnPfx & Coln     ' Pfx
Dim C1$: C1 = P & "Id"           ' Fld-1
Dim C2$: C2 = P & "Cnt"          ' Fld-2
Dim Ix&: Ix = IxEle(D.Fny, Coln) ' Ix-of-Coln

Dim Dy():                         Dy = D.Dy
Dim DcDrs():                       DcDrs = DcDy(Dy, Ix)
Dim DicId  As Dictionary: Stop ' Set DicId = DieyCnt(DcDrs)
Dim DicCnt As Dictionary: Stop 'Set DicCnt = DiCnt(DcDrs)
Dim NCol%:                      NCol = NDcDy(Dy)
Dim ODy():                       ODy = Dy
Dim Dr, R&: For Each Dr In Itr(Dy)
:                     ReDim Preserve Dr(NCol + 1) ' Extend 2 elements
Dim K:            K = Dr(Ix)
           Dr(NCol) = DicId(K)
       Dr(NCol + 1) = DicCnt(K)
             ODy(R) = Dr                          ' <== Put to ODy
                  R = R + 1
Next
Dim Fny$():              Fny = SyAdd(D.Fny, Sy(C1, C2))                  ' New-Fny
            DrsAdd_valid_cnt_2Dc = Drs(Fny, ODy)
End Function
