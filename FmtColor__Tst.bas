Attribute VB_Name = "FmtColor__Tst"
Option Explicit

Sub Tst()
Dim WrkWs As Worksheet
    Set WrkWs = Src.Wrk.Ws
Dim OrgWs As Worksheet
    Set OrgWs = Src.Org.Ws
    
FmtColor_DoRestore WrkWs
FmtColor_DoRestore OrgWs
End Sub
