Attribute VB_Name = "MxDta_Da_Drs_GDrs"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Drs_GDrs."

Function GDrs(D As Drs, Keycc$, Gpcc$) As Drs
'@D     : :Drs ..Keycc ..Gpcc: ! with fields as described in @Keycc @Gpcc
'@Keycc : :CC                  ! #Key-Column-Names# These columns in @D will be returned as first #NKey columns of returned @@Drs
'@Gpcc  : :CC                  ! #Gp-Column-Names# These columns in @D will be grouped as last-field of returned $$Drs @@
':GDrs: :Drs ! #Grouped-Drs# it is a Drs with #NKey + 1 columns
'            ! whose Fny is with first #NKey is from @Keycc
'             ! .    last column name is in format of "Gp-<Gpc1>-<Gpc2>-..-<GpcN>"
'                              !              which Gpc<i> is coming from @Gpcc

Dim KeyDrs As Drs: KeyDrs = DrsSelFf(D, Keycc)
Dim KeyGRxy():    KeyGRxy = DyGp(KeyDrs.Dy)
Dim DistKeyDy() ': Dim KeyGRxy(), KeyDrs As Drs
    Dim WDy(): WDy = KeyDrs.Dy
    Dim WIRxy: For Each WIRxy In Itr(KeyGRxy)
        Dim WRxy&(): WRxy = WIRxy
        PushI DistKeyDy, WDy(WRxy(0))   '<---
    Next
Dim GpDy(): GpDy = DrsSelFf(D, Gpcc).Dy
Dim ODy(): 'Dim KEyGRxy(), GpDy(), DistKeyDy()
    Dim W2IxDistKey%: W2IxDistKey = 0
    Dim W2IRxy: For Each W2IRxy In Itr(KeyGRxy)
        Dim W2IGpDy()
            Erase W2IGpDy
            Dim W2IRix: For Each W2IRix In W2IRxy
                PushI W2IGpDy, GpDy(W2IRix)
            Next
        Dim W2Dr(): W2Dr = DistKeyDy(W2IxDistKey)
                           PushI W2Dr, W2IGpDy
                           PushI ODy, W2Dr '<--
             W2IxDistKey = W2IxDistKey + 1
    Next
Dim OFny$(): 'Dim KeyDrs As Drs
    OFny = KeyDrs.Fny
           PushI OFny, "Gp-" & Replace(Gpcc, " ", "-")
           
GDrs = Drs(OFny, ODy)
'BrwDrs GDrs: Stop
End Function
