Attribute VB_Name = "DtaErCrtErWs"
Option Explicit

Sub DtaEr_DoCrt_ErWs(DtaEr As TDtaErOpt, WrkWs As Worksheet, Optional ErVer As eErWsVer = eErWsVer.eV1ErAndLst)
Dim Wb As Workbook
    Set Wb = WrkWs.Parent

Select Case ErVer
Case eV1ErAndLst:                          DtaErV1_DoCrt_TwoErWs Wb, DtaEr 'DtaEr_Crt_TwoErWs Wb, A_DePush
Case eV2DropDown, eV3SelInEr, eV4SelInSep: DtaErVx_Crt_OneErWs Wb, DtaEr, ErVer
Case Else: Er "Invalid {ErVer}", ErVer
End Select
End Sub
