VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Fct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Sub AlignTwoWb(DtaWb As Workbook)
Dim Wb As Workbook
If Application.Workbooks.Count > 2 Then
    For Each Wb In Application.Workbooks
        If Wb.FullName <> DtaWb.FullName And _
            yWbMassUpd.FullName <> Wb.FullName Then
            Err.Raise 1, , "Please close workbook[" & Wb.Name & "]"
        End If
    Next
End If
If Application.Workbooks.Count <> 2 Then
    Err.Raise 1, , "Program error: there should have 2 workbooks opened, but now[" & Application.Workbooks.Count & "]"
End If
Dim J%
For J = 2 To DtaWb.Windows.Count
    DtaWb.Windows(J).Close
Next
For J = 2 To yWbMassUpd.Windows.Count
    yWbMassUpd.Windows(J).Close
Next
DtaWb.Windows(1).WindowState = xlNormal
yWbMassUpd.Windows(1).WindowState = xlNormal
Windows.Arrange ArrangeStyle:=xlHorizontal
End Sub

Function FldNm_OfChr$(CostGp$, CostEle$, CharName$)
FldNm_OfChr = Fmt_QQ("Char : ? | ? | ?", CostGp, CostEle, CharName)
End Function

Function FldNm_OfCstRmk$(CostGp$, CostEle$)
FldNm_OfCstRmk = Fmt_QQ("Cost Rmk : ? | ?", CostGp, CostEle)
End Function

Function FldNm_OfCstVal$(CostGp$, CostEle$)
FldNm_OfCstVal = Fmt_QQ("Cost : ? | ?", CostGp, CostEle)
End Function

Function SrcSqvAdr$(R&, C)
SrcSqvAdr = Ws_Adr(yWsMassUpd, R + 6, C)
End Function

