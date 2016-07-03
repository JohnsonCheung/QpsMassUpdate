Attribute VB_Name = "Ide_Md"
Option Explicit

Function CurMd() As CodeModule
Set CurMd = Application.VBE.SelectedVBComponent.CodeModule
End Function

Function Md_ByNm(Nm$) As CodeModule
Set Md_ByNm = Application.VBE.ActiveVBProject.VBComponents(Nm).CodeModule
End Function

Function Md_Nm$(Md As CodeModule)
Md_Nm = Md.Parent.Name
End Function

Function Md_Pj(Md As CodeModule) As VBProject
Set Md_Pj = Md.Parent.Collection.Parent
End Function
