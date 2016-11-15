Option Explicit
' Class defining required constants and look-ups.
' Provides functions to check membership of any dictionary defined here.
' SQLite key words source: http://www.sqlite.org/lang_keywords.html
' When creating SQLite tables I want to use column terms compatible with PostgreSQL
'  so added the allowed data types to the end of the lookup dictionary.
' If I need to extend the list of disallowed terms, I will simply add to this dictionary.

'                                           Test driver subroutine
'Private Sub testAllowedValues()
'    Dim sqliteLu As SqliteLookups
'    Dim testValue As String
'    Set sqliteLu = New SqliteLookups
'    testValue = InputBox("Enter a value", "Testing valid identifiers for SQLite")
'    Debug.Print sqliteLu.valueIsAllowed(testValue)
'End Sub


Private m_DisallowedIdentifiersMap As Dictionary
Private m_AllowedDataTypesMap As Dictionary
Private Sub class_initialize()
    Set m_DisallowedIdentifiersMap = New Dictionary
    Set m_AllowedDataTypesMap = New Dictionary
    ' Disallow all key words and the data types that are to be used in SQLite databases.
    With m_DisallowedIdentifiersMap
        .Add "ABORT", "Keyword"
        .Add "ACTION", "Keyword"
        .Add "ADD", "Keyword"
        .Add "AFTER", "Keyword"
        .Add "ALL", "Keyword"
        .Add "ALTER", "Keyword"
        .Add "ANALYZE", "Keyword"
        .Add "AND", "Keyword"
        .Add "AS", "Keyword"
        .Add "ASC", "Keyword"
        .Add "ATTACH", "Keyword"
        .Add "AUTOINCREMENT", "Keyword"
        .Add "BEFORE", "Keyword"
        .Add "BEGIN", "Keyword"
        .Add "BETWEEN", "Keyword"
        .Add "BY", "Keyword"
        .Add "CASCADE", "Keyword"
        .Add "CASE", "Keyword"
        .Add "CAST", "Keyword"
        .Add "CHECK", "Keyword"
        .Add "COLLATE", "Keyword"
        .Add "COLUMN", "Keyword"
        .Add "COMMIT", "Keyword"
        .Add "CONFLICT", "Keyword"
        .Add "CONSTRAINT", "Keyword"
        .Add "CREATE", "Keyword"
        .Add "CROSS", "Keyword"
        .Add "CURRENT_DATE", "Keyword"
        .Add "CURRENT_TIME", "Keyword"
        .Add "CURRENT_TIMESTAMP", "Keyword"
        .Add "DATABASE", "Keyword"
        .Add "DEFAULT", "Keyword"
        .Add "DEFERRABLE", "Keyword"
        .Add "DEFERRED", "Keyword"
        .Add "DELETE", "Keyword"
        .Add "DESC", "Keyword"
        .Add "DETACH", "Keyword"
        .Add "DISTINCT", "Keyword"
        .Add "DROP", "Keyword"
        .Add "EACH", "Keyword"
        .Add "ELSE", "Keyword"
        .Add "END", "Keyword"
        .Add "ESCAPE", "Keyword"
        .Add "EXCEPT", "Keyword"
        .Add "EXCLUSIVE", "Keyword"
        .Add "EXISTS", "Keyword"
        .Add "EXPLAIN", "Keyword"
        .Add "FAIL", "Keyword"
        .Add "FOR", "Keyword"
        .Add "FOREIGN", "Keyword"
        .Add "FROM", "Keyword"
        .Add "FULL", "Keyword"
        .Add "GLOB", "Keyword"
        .Add "GROUP", "Keyword"
        .Add "HAVING", "Keyword"
        .Add "IF", "Keyword"
        .Add "IGNORE", "Keyword"
        .Add "IMMEDIATE", "Keyword"
        .Add "IN", "Keyword"
        .Add "INDEX", "Keyword"
        .Add "INDEXED", "Keyword"
        .Add "INITIALLY", "Keyword"
        .Add "INNER", "Keyword"
        .Add "INSERT", "Keyword"
        .Add "INSTEAD", "Keyword"
        .Add "INTERSECT", "Keyword"
        .Add "INTO", "Keyword"
        .Add "IS", "Keyword"
        .Add "ISNULL", "Keyword"
        .Add "JOIN", "Keyword"
        .Add "KEY", "Keyword"
        .Add "LEFT", "Keyword"
        .Add "LIKE", "Keyword"
        .Add "LIMIT", "Keyword"
        .Add "MATCH", "Keyword"
        .Add "NATURAL", "Keyword"
        .Add "NO", "Keyword"
        .Add "NOT", "Keyword"
        .Add "NOTNULL", "Keyword"
        .Add "NULL", "Keyword"
        .Add "OF", "Keyword"
        .Add "OFFSET", "Keyword"
        .Add "ON", "Keyword"
        .Add "OR", "Keyword"
        .Add "ORDER", "Keyword"
        .Add "OUTER", "Keyword"
        .Add "PLAN", "Keyword"
        .Add "PRAGMA", "Keyword"
        .Add "PRIMARY", "Keyword"
        .Add "QUERY", "Keyword"
        .Add "RAISE", "Keyword"
        .Add "RECURSIVE", "Keyword"
        .Add "REFERENCES", "Keyword"
        .Add "REGEXP", "Keyword"
        .Add "REINDEX", "Keyword"
        .Add "RELEASE", "Keyword"
        .Add "RENAME", "Keyword"
        .Add "REPLACE", "Keyword"
        .Add "RESTRICT", "Keyword"
        .Add "RIGHT", "Keyword"
        .Add "ROLLBACK", "Keyword"
        .Add "ROW", "Keyword"
        .Add "SAVEPOINT", "Keyword"
        .Add "SELECT", "Keyword"
        .Add "SET", "Keyword"
        .Add "TABLE", "Keyword"
        .Add "TEMP", "Keyword"
        .Add "TEMPORARY", "Keyword"
        .Add "THEN", "Keyword"
        .Add "TO", "Keyword"
        .Add "TRANSACTION", "Keyword"
        .Add "TRIGGER", "Keyword"
        .Add "UNION", "Keyword"
        .Add "UNIQUE", "Keyword"
        .Add "UPDATE", "Keyword"
        .Add "USING", "Keyword"
        .Add "VACUUM", "Keyword"
        .Add "VALUES", "Keyword"
        .Add "VIEW", "Keyword"
        .Add "VIRTUAL", "Keyword"
        .Add "WHEN", "Keyword"
        .Add "WHERE", "Keyword"
        .Add "WITH", "Keyword"
        .Add "WITHOUT", "Keyword"
        .Add "TEXT", "Data type"
        .Add "NUMERIC", "Data type"
        .Add "INTEGER", "Data type"
        .Add "DATE", "Data type"
    End With
    ' Not interested in the values, just want to use the keys for a membership test.
    With m_AllowedDataTypesMap
        .Add "TEXT", 1
        .Add "NUMERIC", 1
        .Add "INTEGER", 1
        .Add "DATE", 1
    End With
End Sub

' Check if the given value is in the m_DisallowedIdentifiersMap.
' Text case and leading and trailing white space are ignored.
Public Function valueIsAllowed(ByVal value) As Boolean
    value = UCase(Trim(value))
    If m_DisallowedIdentifiersMap.Exists(value) Then
        valueIsAllowed = False
        Exit Function
    End If
    valueIsAllowed = True
End Function
' Use a regular expression to check the given identifier:
'  1: Begins with a letter or underscore
'  2: The rest of the identifier contains word characters (letters, numbers or underscore).
' Ensure "Tools -> References -> Microsoft VBScript Regular Expressions 5.5" is set.
Public Function isValidIdentifier(identifier) As Boolean
    Dim regexPat As String
    Dim regex As RegExp
    Dim i As Integer
    Set regex = New RegExp
    regexPat = "^[a-zA-Z_]\w+$"
    regex.Pattern = regexPat
    If regex.Test(identifier) Then
        isValidIdentifier = True
    Else
        isValidIdentifier = False
    End If
End Function
' Run two checks on any given identifier:
'  1: It is not an SQLite reserved word or specified column data type
'  2: It is a valid name that passes a prescribed pattern
Public Function CanUseIdentifier(identifier As String) As Boolean
    If valueIsAllowed(identifier) And isValidIdentifier(identifier) Then
        CanUseIdentifier = True
    Else
        CanUseIdentifier = False
    End If
End Function
' Checks if a given data type is allowed by checking a dictionary lookup.
' Leading and trailing whitespace and text case are ignored in the comparison.
Public Function CanUseDataType(datatype As String) As Boolean
    datatype = UCase(Trim(datatype))
    If m_AllowedDataTypesMap.Exists(datatype) Then
        CanUseDataType = True
    Else
        CanUseDataType = False
    End If
End Function
