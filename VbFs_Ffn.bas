Attribute VB_Name = "VbFs_Ffn"
Option Explicit
Type FInfo
    Ffn As String
    Sz As Long
    Tim As Date
End Type

Function Ffn_AddSfx$(Ffn, Sfx$)
Ffn_AddSfx = CutExt(Ffn) & Sfx & Ffn_Ext(Ffn)
End Function

Function Ffn_CutExt$(Ffn)
Ffn_CutExt = TakBefLastX(Ffn, ".")
End Function

Sub Ffn_DltIfExist(Ffn$)
If Ffn_IsExist(Ffn) Then Kill Ffn
End Sub

Function Ffn_Ext$(Ffn)
Ffn_Ext = TakAftLastX(Ffn, ".", AlsoTakX:=True)
End Function

Function Ffn_FInfo(Ffn$) As FInfo
Dim O As FInfo
With O
    .Ffn = Ffn
    .Sz = FileLen(Ffn)
    .Tim = FileDateTime(Ffn)
End With
Ffn_FInfo = O
End Function

Function Ffn_ReplPth$(Ffn$, Pth$)
Ffn_ReplPth = Pth & Ffn_Fn(Ffn)
End Function

Function Ffn_Fn$(Ffn)
Ffn_Fn = Dft(TakAftLastX(Ffn, "\"), Ffn)
End Function

Function Ffn_Fnn$(Ffn)
Ffn_Fnn = CutExt(Ffn_Fn(Ffn))
End Function

Function Ffn_IsExist(Ffn$) As Boolean
Ffn_IsExist = Dir(Ffn) <> ""
End Function

Function Ffn_Pth$(Ffn)
Ffn_Pth = TakBefLastX(Ffn, "\", AlsoTakX:=True)
End Function

Function Ffn_ReplExt$(Ffn, NewExt$)
If ChrBeg(NewExt) <> "." Then Err.Raise 1
Ffn_ReplExt = CutExt(Ffn) & NewExt
End Function

Function Ffn_ReplRelativePth$(Ffn, RelativePth$)
If ChrEnd(RelativePth) <> "\" Then Err.Raise 1
Ffn_ReplRelativePth = Ffn_Pth(Ffn) & RelativePth & Ffn_Fn(Ffn)
End Function
