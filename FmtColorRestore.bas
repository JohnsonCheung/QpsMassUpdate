Attribute VB_Name = "FmtColorRestore"
Option Explicit
Private Type B
    R1 As Long
    R2 As Long
    Row5HSqv As Variant
    DtaBdyRge As Range
End Type
Private A_Ws As Worksheet
Private B As B

Sub FmtColor_DoRestore(WrkWs As Worksheet)
Set A_Ws = WrkWs

Dim LastCno%
    LastCno = Ws_RC(A_Ws, 5, 1).End(xlToRight).Column

Set B.DtaBdyRge = A_Ws.ListObjects(1).DataBodyRange
B.Row5HSqv = Ws_RCC(A_Ws, 5, 1, LastCno).Value
B.R1 = B.DtaBdyRge.Row
B.R2 = B.R1 + ZDtaBdyRge.Rows.Count - 1 ' "-1" is to exclude the total row
Const FillerColor = 10092543   ' LightYellow
Const TotColColor = 13238235   ' LightGreen
ZDo_Clear_Color
ZDo_Set_OneColor ZFillerCnoAy, 10092543   ' LightYellow
ZDo_Set_OneColor ZTotColCnoAy, 13238235   ' LightGreen
End Sub

Private Sub ZDo_Clear_Color()
Dim D As Range
    Set D = ZDtaBdyRge
With Ws_RR(D.Worksheet, 1, 5)
    .Interior.ColorIndex = xlAutomatic ' SOme error will highlight the header row
    .Interior.Pattern = XlPattern.xlPatternAutomatic
    .Font.ColorIndex = xlAutomatic
    .VerticalAlignment = xlVAlignTop
End With
With D
    .Font.ColorIndex = xlAutomatic
    .Interior.ColorIndex = xlAutomatic
    .Interior.Pattern = XlPattern.xlPatternAutomatic
    .VerticalAlignment = xlVAlignTop
End With
End Sub

Private Sub ZDo_Set_OneColor(Cno%(), Color&)
Dim J%
For J = 0 To UBound(Cno)
    Dim R%
    For R = 1 To 5
        If R = 3 Then GoTo Nxt
        Ws_RC(A_Ws, R, Cno(J)).Interior.Color = Color
Nxt:
    Next
    Ws_CRR(A_Ws, Cno(J), ZR1, ZR2).Interior.Color = Color
Next
End Sub

Private Property Get ZDtaBdyRge() As Range
Set ZDtaBdyRge = B.DtaBdyRge
End Property

Private Property Get ZFillerCnoAy() As Integer()
Dim HSqv
    HSqv = ZRow5HSqv
Dim O%()
    Dim J%
    For J = 1 To UBound(HSqv, 2)
        If HSqv(1, J) Like "ChrGp??Filler" Then
            Push O, J
        End If
    Next
ZFillerCnoAy = O
End Property

Private Property Get ZR1&()
ZR1 = B.R1
End Property

Private Property Get ZR2&()
ZR2 = B.R2
End Property

Private Property Get ZRow5HSqv()
ZRow5HSqv = B.Row5HSqv
End Property

Private Property Get ZTotColCnoAy() As Integer()
Dim HSqv
    HSqv = ZRow5HSqv
Dim O%()
    Dim J%
    For J = 1 To UBound(HSqv, 2)
        If HSqv(1, J) Like "*Tot" Or HSqv(1, J) = "SkuCost" Then
            Push O, J
        End If
    Next
ZTotColCnoAy = O
End Property
