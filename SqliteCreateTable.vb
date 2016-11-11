Option Explicit
' Class to encapsulate table creation
Private m_sqliteConn As ADODB.connection
Private Sub class_initialize()

End Sub
Public Function SetSqliteConnection(sqliteConn As ADODB.connection) As Boolean
    Set m_sqliteConn = sqliteConn
    SetSqliteConnection = True
End Function
Public Function TableExists(tableName As String) As Boolean
    Dim sql As String
    Dim rs As ADODB.Recordset
    sql = "SELECT COUNT(*) table_count FROM sqlite_master WHERE type = 'table' AND tbl_name = '" & tableName & "'"
    Set rs = m_sqliteConn.Execute(sql)
    rs.MoveFirst
    MsgBox rs.Fields("table_count").value
End Function
Public Function ExecuteCreateTable(ddl As String) As Boolean

End Function
