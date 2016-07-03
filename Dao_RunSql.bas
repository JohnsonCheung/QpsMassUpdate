Attribute VB_Name = "Dao_RunSql"
Option Explicit

Function RunSql_IsAny(Db As Database, Sql$)
With Db.OpenRecordset(Sql)
    RunSql_IsAny = Not .EOF
    .Close
End With
End Function

Function RunSql_Val(Db As Database, Sql$)
With Db.OpenRecordset(Sql)
    RunSql_Val = .Fields(0).Value
    .Close
End With
End Function
