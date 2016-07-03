Attribute VB_Name = "Globals"
Option Explicit
Public Const LstWsNm = "ChrList"
Public Const ErWsNm = "Error"
Public Const DtaChgWsNm = "DataChanged"
Public Const LnkWsNm = "Error"
Public Const WrkWsNm = "Working"
Public Const OrgWsNm = "Original"
Public Fct As New Fct
Public Enm As New Enm
Public QErTxt As New QErTxt
Public ChrDefInf As New CfgChrDefFmPgmFx
Type TSrc
    Wrk As TWsInf
    Org As TWsInf
End Type

Sub Cmd_1Opn_ImpFdr()
Pth_Opn Fs_ImpFdr
End Sub

Sub Cmd_1Vdt_V1ErAndLst()
Do_Vdt Src, eV1ErAndLst
End Sub

Sub Cmd_1Vdt_V2DropDown()
Do_Vdt Src, eV2DropDown
End Sub

Sub Cmd_1Vdt_V3SelInEr()
Do_Vdt Src, eV3SelInEr
End Sub

Sub Cmd_1Vdt_V4SelInSep()
Do_Vdt Src, eV4SelInSep
End Sub

Sub Cmd_2Import()
Do_Import
End Sub

Sub Cmd_2Opn_FbFdr()
Pth_Opn Fs_FbFdr
End Sub

Sub Cmd_4Cpy_CfgFx()
Do_CfgCpy
End Sub

Sub Cmd_4Opn_CfgFdr()
Pth_Opn Fs_CfgFdr: Exit Sub
End Sub

Sub Cmd_4Opn_CfgFx()
Workbooks.Open Fs_CfgFx
End Sub

Sub Cmd_DltLastQDtePj()
Do_Dlt_LastQDtaPj
End Sub

Property Get EvtHandlerObjDic() As Dictionary
Static X As New Dictionary
Set EvtHandlerObjDic = X
End Property

Sub RenMd(Pfx$, ToPfx$)
PjNm_RenMd "MassUpd", Pfx, ToPfx
End Sub

Private Sub ChrDefInf__Tst()
Dim A As CfgChrDefFmPgmFx
    Set A = ChrDefInf
Dim B$()
B = A.VdtChrCdAy
Stop
End Sub
