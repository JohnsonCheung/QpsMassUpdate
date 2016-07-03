Attribute VB_Name = "Dao_DbTInf"
Option Explicit

Type DbT
    D As Database
    T As String
End Type

Function DbT(D As Database, T$) As DbT
Dim O As DbT
    Set O.D = D
    O.T = T
DbT = O
End Function
