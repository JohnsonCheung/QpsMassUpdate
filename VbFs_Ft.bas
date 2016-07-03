Attribute VB_Name = "VbFs_Ft"
Option Explicit

Sub Ft_Brw(Ft$)
Shell Fmt_QQ("NotePad ""?""", Ft)
End Sub

Sub Ft_DltIfExist(Ft$)

End Sub

Sub Ft_WrtStr(Ft$, S)
Dim F%
    F = FreeFile(1)
Open Ft For Output As #F
Print #F, S
Close #F
End Sub
