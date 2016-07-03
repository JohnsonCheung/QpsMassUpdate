Attribute VB_Name = "MacroErWsChgV3SelInEr__Tst"
Option Explicit


Private A_ErVer
Private A_ErWs As Worksheet

Private Sub ZCrt_ErWs_Macro()
A_ErVer = eV3SelInEr
Set A_ErWs = ErWsV3

Dim Fn$, Fn1$
    Select Case A_ErVer
    Case eV2DropDown: Fn = "ErWs_WsChg_V2DropDown"
    Case eV3SelInEr:  Fn = "ErWs_WsChg_V3SelInEr":   Fn1 = "ErWs_WsSelChg_V3SelInEr"
    Case eV4SelInSep: Fn = "ErWs_WsChg_eV4SelInSep": Fn1 = "ErWs_WsSelChg_V4SelInSep"
    Case Else:        Er "Given {ErVer} should be V2..V4", A_ErVer
    End Select
Const Evt1$ = "Change"
Const Evt2$ = "SelectionChange"
Ws_Clr_Md A_ErWs
Ws_Crt_EvtMth_CallingFn A_ErWs, Evt1, "jjMassUpd", , Fn        '<== Add Worksheet_change
If Fn1 = "" Then Exit Sub
Ws_Crt_EvtMth_CallingFn A_ErWs, Evt2, "jjMassUpd", , Fn1       '<== Optional Add Worksheet_Selectionchange
End Sub
