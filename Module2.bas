Attribute VB_Name = "Module2"
Option Explicit
Function CmpTy_IsWithMd(Ty As vbext_ComponentType) As Boolean
Select Case Ty
Case _
    vbext_ComponentType.vbext_ct_ClassModule, _
    vbext_ComponentType.vbext_ct_StdModule, _
    vbext_ComponentType.vbext_ct_Document
    CmpTy_IsWithMd = True
End Select
End Function

Function CmpTy_Ext$(Ty As vbext_ComponentType)
Dim O$
Select Case Ty
Case vbext_ComponentType.vbext_ct_ClassModule: O = ".cls"
Case vbext_ComponentType.vbext_ct_StdModule: O = ".bas"
Case vbext_ComponentType.vbext_ct_Document: O = ".cls"
Case Else: Er "{CmpTy} invalid.  Valid: Class/Std/Doc", Ty
End Select
CmpTy_Ext = O
End Function
