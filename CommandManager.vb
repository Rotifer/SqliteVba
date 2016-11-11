Option Explicit
' Class to wrap the Command class
Private Type TypeMappings
    BooleanMap As ADODB.DataTypeEnum
    ByteMap As ADODB.DataTypeEnum
    CurrencyMap As ADODB.DataTypeEnum
    DateMap As ADODB.DataTypeEnum
    DoubleMap As ADODB.DataTypeEnum
    IntegerMap As ADODB.DataTypeEnum
    LongMap As ADODB.DataTypeEnum
    SingleMap As ADODB.DataTypeEnum
    StringMap As ADODB.DataTypeEnum
End Type
Private mappings As TypeMappings
Private m_cmd As ADODB.Command
Private m_sqlText As String
Private Sub class_initialize()
    mappings.BooleanMap = adBoolean
    mappings.ByteMap = adInteger
    mappings.CurrencyMap = adCurrency
    mappings.DateMap = adDate
    mappings.DoubleMap = adDouble
    mappings.IntegerMap = adInteger
    mappings.LongMap = adInteger
    mappings.SingleMap = adSingle
    mappings.StringMap = adVarChar
End Sub
Public Property Get Command() As ADODB.Command
    Set Command = m_cmd
End Property

Public Property Let sql(sqlText As String)
    m_sqlText = sqlText
End Property
Public Function PrepareCommand(conn As ADODB.connection, sqlText As String) As Boolean
    Set m_cmd = New ADODB.Command
    m_cmd.ActiveConnection = conn
    m_cmd.CommandText = sqlText
    PrepareCommand = True
End Function
Public Function SetStringParameter(value As String, direction As ADODB.ParameterDirectionEnum) As Boolean
    Dim param As ADODB.Parameter
    Set param = m_cmd.CreateParameter
    param.Type = mappings.StringMap
    param.direction = direction
    param.value = value
    param.Size = 100
    m_cmd.Parameters.Append param
    SetStringParameter = True
End Function
