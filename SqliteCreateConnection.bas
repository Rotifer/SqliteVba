Option Explicit
' Class to wrap creation of a Connection instance for an sQLite database
Private Const m_connStrTemplate As String = "DRIVER=SQLite3 ODBC Driver;Database=[PATH_TO_DATABASE_FILE];LongNames=0;Timeout=1000;NoTXN=0;SyncPragma=NORMAL;StepAPI=0;FKSupport=true"
Private m_connection As ADODB.Connection
Private m_lastError As String
Private Const NOERRORSFLAG As String = "None"
Private m_fileChkr As FileChecker
Private Sub class_initialize()
    Set m_fileChkr = New FileChecker
    m_lastError = NOERRORSFLAG
End Sub

Public Function GetConnection(dbPath As String) As ADODB.Connection
    Dim fullConnStr As String
    If Not m_fileChkr.PathIsOk(dbPath) Then
        m_lastError = "The path given does not exists"
        Exit Function
    End If
    fullConnStr = Replace(m_connStrTemplate, "[PATH_TO_DATABASE_FILE]", dbPath)
    Set m_connection = New ADODB.Connection
    m_connection.Open fullConnStr
    Set GetConnection = m_connection
End Function

Public Property Get LastError() As String
    LastError = m_lastError
End Property
Private Sub class_terminate()

End Sub
