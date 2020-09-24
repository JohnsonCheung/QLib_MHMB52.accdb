Attribute VB_Name = "MxOfc_P_Kill"
Option Compare Text
Const CMod$ = "MxOfc_P_Kill."
Option Explicit

Sub KillAcs()
Dim AcsGet As Access.Application: Set AcsGet = WAcsGet
Dim ObjPtrGet&: ObjPtrGet = ObjPtr(AcsGet)
Dim ObjPtrCur&: ObjPtrCur = ObjPtr(Acs)
While ObjPtr(AcsGet) <> 0
    Dim J%: ThwLoopTooMuch CSub, J, 100
    If ObjPtrCur = ObjPtr(AcsGet) Then Exit Sub
    QuitAcs AcsGet
    Set AcsGet = WAcsGet
Wend
End Sub
Private Function WAcsGet() As Access.Application: Set WAcsGet = GetObject(, "Access.Application"): End Function

Sub KillXls()
Dim XlsGet As Excel.Application: Set XlsGet = WXlsGet
Dim ObjPtrGet&: ObjPtrGet = ObjPtr(XlsGet)
Dim ObjPtrCur&: ObjPtrCur = ObjPtr(Xls)
While ObjPtr(XlsGet) <> 0
    If ObjPtrCur = ObjPtr(XlsGet) Then Exit Sub
    QuitXls XlsGet
    Set XlsGet = WXlsGet
Wend
End Sub
Private Function WXlsGet() As Excel.Application: Set WXlsGet = GetObject(, "Excel.Application"): End Function
