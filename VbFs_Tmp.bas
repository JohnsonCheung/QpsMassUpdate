Attribute VB_Name = "VbFs_Tmp"
Option Explicit
Const CurMdNm$ = "FctFs"

Function Tmp_Ffn$(Ext$, Optional Nm$)
Tmp_Ffn = Tmp_Root & TimStmp & Ext
End Function

Function Tmp_Ft$(Optional Nm$)
Tmp_Ft = Tmp_Ffn(".txt", Nm)
End Function

Function Tmp_Pth$()
Tmp_Pth = Tmp_Root & TimStmp & "\"
End Function

Function Tmp_Root$()
Tmp_Root = "C:\Users\User\Temp\"
End Function
