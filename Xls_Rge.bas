Attribute VB_Name = "Xls_Rge"
Option Explicit

Function Rge_C(Rge As Range, C) As Range
Set Rge_C = Rge_CRR(Rge, C, 1, Rge.Rows.Count)
End Function

Function Rge_CRR(Rge As Range, C, R1, R2) As Range
Set Rge_CRR = Rge.Worksheet.Range(Rge.Cells(R1, C), Rge.Cells(R2, C))
End Function

Function Rge_R1R2(Rge As Range) As TR1R2
Dim R1&, R2&
    R1 = Rge.Row
    R2 = R1 + Rge.Rows.Count - 1
Dim O As TR1R2
    O.R1 = R1
    O.R2 = R2
Rge_R1R2 = O
End Function

Function Rge_RC(Rge As Range, R, C) As Range
Set Rge_RC = Rge.Cells(R, C)
End Function

Function Rge_RCC(Rge As Range, R, C1, C2) As Range
Set Rge_RCC = Rge.Worksheet.Range(Rge.Cells(R, C1), Rge.Cells(R, C2))
End Function

Function Rge_RCRC(Rge As Range, R1, C1, R2, C2) As Range
Set Rge_RCRC = Rge.Worksheet.Range(Rge_RC(Rge, R1, C1), Rge_RC(Rge, R2, C2))
End Function

Function Rge_RR(Rge As Range, R1, R2) As Range
Set Rge_RR = Rge_CRR(Rge, 1, R1, R2).EntireRow
End Function

Function Rge_Ws(Rge As Range) As Worksheet
Set Rge_Ws = Rge.Parent
End Function
