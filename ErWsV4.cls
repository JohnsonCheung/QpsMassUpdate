VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErWsV4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
jjMassUpd.ErWs_WsChg_V4SelInSep Target
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
jjMassUpd.ErWs_WsSelChg_V4SelInSep Target
End Sub
