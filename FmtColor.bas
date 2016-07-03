Attribute VB_Name = "FmtColor"
Option Explicit

Sub FmtColor_DoPaint_Red(OrgWs As Worksheet, WrkWs As Worksheet, Adr As RedAdr)
Dim J&
For J = 0 To UB(Adr.Adr)
    Dim Ws As Worksheet
        Select Case Adr.WhichWs(J)
        Case eOrgWs: Set Ws = OrgWs
        Case eWrkWs: Set Ws = WrkWs
        Case Else: Er "Invalid {WhichWs}", Adr.WhichWs(J)
        End Select
    Ws.Range(Adr.Adr(J)).Interior.Color = rgbRed
Next
End Sub

Sub FmtColor_DoPaint_Yellow(WrkWs As Worksheet, Adr As YellowAdr)
Dim J&
With Adr
    For J = 0 To UB(Adr.C)
        Ws_RC(WrkWs, .R(J), .C(J)).Interior.Color = rgbYellow
    Next
End With
End Sub
