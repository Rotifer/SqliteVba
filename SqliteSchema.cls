Option Explicit
' SqliteSchema is used to check the database given to it in the connection instance for
' objects such as tables and views.
Private m_sqliteConn As ADODB.Connection
Private m_recsetProcessor As RecordsetProcessor
Private m_cmdMgr As CommandManager
Private Sub class_initialize()

End Sub

Public Function Setup(sqliteConn As ADODB.Connection) As Boolean
    Set m_sqliteConn = sqliteConn
    Set m_cmdMgr = New CommandManager
    Set m_recsetProcessor = New RecordsetProcessor
    Setup = True
End Function
' Query the SQLite metadata table called "sqlite_master" to determine if a database object of
'  a given type and name exists.
' Used to check if tables or views with a given name already exists in the database.
Private Function databaseObjectExists(objectType As String, objectName As String) As Boolean
    Dim SQL As String
    Dim rowCount As Integer
    Dim rs As ADODB.Recordset
    SQL = "SELECT COUNT(*) object_count FROM sqlite_master WHERE type = ? AND tbl_name = ?"
    m_cmdMgr.Setup m_sqliteConn, SQL
    m_cmdMgr.SetStringParameter objectType, adParamInput
    m_cmdMgr.SetStringParameter objectName, adParamInput
    Set rs = m_cmdMgr.Command.Execute
    m_recsetProcessor.Setup rs
    rowCount = m_recsetProcessor.GetSingleValue("object_count")
    If rowCount = 0 Then
        databaseObjectExists = False
        Exit Function
    End If
    databaseObjectExists = True
End Function
Public Function TableExists(tableName As String) As Boolean
    TableExists = databaseObjectExists("table", tableName)
End Function

Public Function ViewExists(viewname As String) As Boolean
    TableExists = databaseObjectExists("view", viewname)
End Function
