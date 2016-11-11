# SqliteCreateConnection

Create a connection to an SQLite database using the class *SqliteCreateConnection*.

```
Private Sub CreateConnection()
    Dim createConn As SqliteCreateConnection
    Dim dbPath As String
    Dim conn As ADODB.connection
    Set createConn = New SqliteCreateConnection
    dbPath = "<windows path here"
    Set conn = createConn.GetConnection(dbPath)
    If Not conn Is Nothing Then
        conn.Close
    End If
    Debug.Print createConn.LastError
End Sub
```

# FileChecker

Try out this class with real and fictitious paths for files.

```
Private Sub RunTests()
    Dim fullPath As String
    Dim fileChkr As FileChecker
    Set fileChkr = New FileChecker
    fullPath = "L:\Lab\license.pdf"
    Debug.Print fileChkr.DirectoryExists(fullPath)
    Debug.Print fileChkr.GetBasename(fullPath)
    Debug.Print fileChkr.GetDirname(fullPath)
    Debug.Print fileChkr.PathIsOk(fullPath)
End Sub
```

Will add driver subs for other classes soon.
