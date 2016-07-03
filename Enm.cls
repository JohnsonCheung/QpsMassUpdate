VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Enm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function R1FormulaChkColStr$(ChkColas As eR1FormulaChkCol)
Stop
End Function

Function WhichWs(WsNm$) As eWhichWs
Dim O$
Select Case WsNm
Case OrgWsNm: O = eWhichWs.eOrgWs
Case WrkWsNm: O = eWhichWs.eWrkWs
Case Else: Er "Invalid {WsNm}", WsNm
End Select
WhichWs = O
End Function

Function WhichWsNm$(W As eWhichWs)
Dim O$
Select Case W
Case eWhichWs.eOrgWs: O = OrgWsNm
Case eWhichWs.eWrkWs: O = WrkWsNm
Case Else: Er "Invalid {WhichWs}", W
End Select
WhichWsNm = O
End Function
