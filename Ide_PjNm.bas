Attribute VB_Name = "Ide_PjNm"
Option Explicit

Sub PjNm_CpyMd(PjNm$, ToPjNm$)
Dim Pj As VBProject
Dim ToPj As VBProject
Set Pj = Pj_ByNm(PjNm)
Set ToPj = Pj_ByNm(ToPjNm)
Dim A$(), J%, Md As CodeModule
A = Pj_MdNmAy(Pj)
For J = 0 To UB(A)
'    Md_CpyToPj Md, ToPj
Next
End Sub

Sub PjNm_RenMd(PjNm$, Pfx$, ToPfx$)
Pj_RenMd Pj_ByNm(PjNm), Pfx, ToPfx
End Sub
