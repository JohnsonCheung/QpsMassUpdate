Attribute VB_Name = "Ide_Pj"
Option Explicit

Function CurPj() As VBProject
Set CurPj = Application.VBE.ActiveVBProject
End Function

Function Pj_ByNm(Nm$) As VBProject
Dim Pj As VBProject
For Each Pj In Application.VBE.VBProjects
    If Pj.Name = Nm Then Set Pj_ByNm = Pj
Next
End Function

Function Pj_IsCmp(Pj As VBProject, CmpNm$) As Boolean
Dim C As VBComponent
For Each C In Pj.VBComponents
    If C.Name = CmpNm Then Pj_IsCmp = True: Exit Function
Next
End Function

Function Pj_MdNmAy(Pj As VBProject) As String()
Dim O$()
Dim C As VBComponent
For Each C In Pj.VBComponents
    If C.Type = vbext_ct_StdModule Then
        Push O, C.Name
    End If
Next
Pj_MdNmAy = O
End Function
Function Pj_SrcPth$(Pj As VBProject)
Dim Fdr$, PjFn$, PjFfn$
PjFfn = Pj.Filename
PjFn = Ffn_Fn(PjFfn)
Pj_SrcPth = Ffn_Pth(PjFfn) & "Src\" & PjFn & "\"
End Function

Sub Pj_RenMd(Pj As VBProject, Pfx$, ToPfx$)
Dim Ay$(), J%, Nm$, NewNm$
Ay = Ay_SubSet_ByPfx(Pj_MdNmAy(Pj), Pfx)
For J = 0 To UB(Ay)
    Nm = Ay(J)
    NewNm = Str_ReplPfx(Nm, Pfx, ToPfx)
    If Pj_IsCmp(Pj, NewNm) Then
        Debug.Print NewNm; "<== Exist"
    Else
        Pj.VBComponents(Ay(J)).Name = NewNm
    End If
Next
End Sub
