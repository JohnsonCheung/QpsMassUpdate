Attribute VB_Name = "Dao_PrpInf"
Option Explicit

Type Prp
    T As TableDef
    FldNm As String
    PrpNm As String
End Type

Function Prp(T As TableDef, FldNm$, PrpNm$) As Prp
Dim O As Prp
Set O.T = T
O.FldNm = FldNm
O.PrpNm = PrpNm
Prp = O
End Function

Sub Prp_Add(P As Prp, V)
Dim T As Dao.DataTypeEnum
    T = Var_DaoTy(V)
Dim F As Field
    Set F = P.T.Fields(P.FldNm)
Dim A As Dao.Property
    Set A = F.CreateProperty(P.PrpNm, T, V)
P.T.Fields(P.FldNm).Properties.Append A
End Sub

Function Prp_Exist(P As Prp) As Boolean
Dim I As Dao.Property
For Each I In P.T.Fields(P.FldNm).Properties
    If I.Name = P.PrpNm Then Prp_Exist = True: Exit Function
Next
End Function

Property Get Prp_Val(P As Prp)
If Prp_Exist(P) Then Prp_Val = P.T.Fields(P.FldNm).Properties(P.PrpNm).Value
End Property

Property Let Prp_Val(P As Prp, V)
If Prp_Exist(P) Then
    P.T.Fields(P.FldNm).Properties(P.PrpNm).Value = V
Else
    Prp_Add P, V
End If
End Property

Sub Prp_Val__Tst()
Dim D As Database
    Set D = OpenDatabase("C:\Temp\a.accdb")
Dim P As Prp
    Set P.T = D.TableDefs("Table1")
    P.FldNm = "Fk1"
    P.PrpNm = "Ele"
Debug.Assert Prp_Exist(P) = True
Prp_Val(P) = "Id"
Debug.Assert Prp_Exist(P) = True
Dim Act$
    Act = Prp_Val(P)
Debug.Assert Act = "Id"
End Sub

Function Var_DaoTy(V) As Dao.DataTypeEnum
Var_DaoTy = VbTy_DaoTy(VarType(V))
End Function

Function VbTy_DaoTy(VbTy As VbVarType) As Dao.DataTypeEnum
Dim O As Dao.DataTypeEnum
Select Case VbTy
Case VbVarType.vbString: O = dbText
Case Else: Stop
End Select
VbTy_DaoTy = O
End Function
