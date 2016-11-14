Option Explicit
' Class to encapsulate table creation
' Uses SqliteSchema to check what is already defined in the database
' Uses SqliteLookups to check if arguments for methods are valid.
Private m_sqliteConn As ADODB.Connection
Private m_schema As SqliteSchema
Private m_isSetup As Boolean
Private m_sqliteLu As SqliteLookups
' An array to hold the table name, the column definitions and the constraints as they are added
'  in the methods below.
Private m_tableColumns() As String
Private m_tableNameHasBeenSet As Boolean
Private m_AtleastOneColumnAdded As Boolean
Private m_tableDefinitionHeader As String
Private m_tableName As String
' Set two flags to ensure the table name has been set and atleast one column has been added.
Private Sub class_initialize()
    'Set some flags
    m_tableNameHasBeenSet = False
    m_AtleastOneColumnAdded = False
End Sub
' Requires a SQLite connection to set up the class
' Uses instances of SqliteSchema and SqliteLookups to run various validity checks.
Public Function Setup(sqliteConn As ADODB.Connection) As Boolean
    Set m_sqliteConn = sqliteConn
    Set m_schema = New SqliteSchema
    Set m_sqliteLu = New SqliteLookups
    m_schema.Setup m_sqliteConn
    m_isSetup = True
    Setup = True
End Function

' Set the name to be used for the table.
' Sets member m_tableDefinitionHeader.
' Checks if the table name has already been set and, if true, exits.
' Checks if the given table name is valid, if not returns a string saying it is invalid.
' A return value of the given table name indicates success, any other value indicates failure.
Public Function AddTableName(tableName As String) As String
    If m_tableNameHasBeenSet Then
        AddTableName = "Table name has already been set!"
        Exit Function
    End If
    If m_sqliteLu.CanUseIdentifier(tableName) Then
        m_tableDefinitionHeader = "CREATE TABLE " & tableName & "(" & vbCrLf
        m_tableNameHasBeenSet = True
        AddTableName = tableName
        m_tableName = tableName
    Else
        AddTableName = "invalid table name!"
    End If
End Function
' Convenience function to add the SQLite auto-incrementing database-generated primary key column
' Column addition is delegated to AddColumn
Public Function AddAutoIncrementPrimaryKeyColumn(columnName As String) As Integer
    Dim columnType As String
    Dim columnConstraint As String
    columnType = "INTEGER"
    columnConstraint = "PRIMARY KEY"
    AddAutoIncrementPrimaryKeyColumn = AddColumn(columnName, columnType, columnConstraint)
End Function

' Adds a column definition to the member array m_tableColumns.
' Checks that the column name is valid and that a recognized column type has been given.
' Currently it does not check for validity of the constraint.
Public Function AddColumn(columnName As String, columnType As String, Optional columnConstraint As String = "") As Boolean
    Dim columnElementToAdd As String
    Dim elementAdded As Boolean
    Dim nextColumnIdx As Integer
    If Not m_AtleastOneColumnAdded Then
        ReDim Preserve m_tableColumns(0)
        nextColumnIdx = 0
    Else
        nextColumnIdx = UBound(m_tableColumns) + 1
        ReDim Preserve m_tableColumns(nextColumnIdx)
    End If
    If m_sqliteLu.isValidIdentifier(columnName) And m_sqliteLu.CanUseDataType(columnType) Then
        ' Make the string and prepend two spaces to make the final generated DDL statement more readable
        columnElementToAdd = "  " & columnName & " " & UCase(columnType) & " " & columnConstraint
        ' RTRIM here is to remove trailing space if no columnConstraint s given.
        m_tableColumns(nextColumnIdx) = RTrim(columnElementToAdd)
        m_AtleastOneColumnAdded = True
        AddColumn = elementAdded
    Else
        AddColumn = False
    End If
End Function
' To be called after the table name has beeb set and all columns have been added.
' It adds the closing ")".
' It will not create the valid DDL if no columns have been added or if no table name has been set
Public Function GenerateFinalCreateTableDdl() As String
    Dim finalCreateTableDdl As String
    If m_tableNameHasBeenSet And m_AtleastOneColumnAdded Then
        finalCreateTableDdl = m_tableDefinitionHeader & Join(m_tableColumns, "," & vbCrLf) & ")"
        GenerateFinalCreateTableDdl = finalCreateTableDdl
    Else
        GenerateFinalCreateTableDdl = "Cannot generate final DDL. Either no table name given or no columns specified"
    End If
End Function
' To be used when the SQL to create the table is fully formed and valid.
' Return True if the DDL executes successfully
' It returns False when the table name already exists.
Public Function ExecuteCreateTable(createTableDdl As String) As Boolean
    If Not m_schema.TableExists(m_tableName) Then
        m_sqliteConn.Execute createTableDdl
        ExecuteCreateTable = True
    Else
        ExecuteCreateTable = False
    End If
End Function
