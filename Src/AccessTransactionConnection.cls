VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccessTransactionConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Frabasic.DataAccess")
Option Explicit

Private Const IS_CONNECTED = 1

Public Connection As ADODB.Connection
Public RecordSet As ADODB.RecordSet
Public Command As ADODB.Command
Public Sql As StringBuilder

Private Sub Class_Initialize()
    Set Connection = New ADODB.Connection
    Let Connection.CursorLocation = adUseClient
    Set Command = New ADODB.Command
    Set RecordSet = New ADODB.RecordSet
    Set Sql = New StringBuilder
End Sub

Public Function Begin() As ADODB.Connection
    
    If Connection.State <> IS_CONNECTED Then
        Connection.Open AppProperty.ConnectionString: Connection.BeginTrans
    End If
    
    Set Begin = Connection
    
End Function

Public Function Commit() As Boolean
    
    On Error GoTo Exception
    
    Call Connection.CommitTrans: Let Commit = True
    
    Exit Function

Exception:
    
    Let Commit = False

End Function

Public Function ExecuteReader(ByVal SqlCommand As String) As ADODB.RecordSet
    
    Command.ActiveConnection = Begin
    Command.CommandType = adCmdText
    Command.CommandText = SqlCommand
      
    Set RecordSet = Command.Execute
    Set ExecuteReader = RecordSet
    
    ClearParameter
    Sql.Clear
    
End Function

Public Function ExecuteScalar(ByVal SqlCommand As String) As Variant
    
    Command.ActiveConnection = Connection
    Command.CommandType = adCmdText
    Command.CommandText = SqlCommand
      
    Set RecordSet = Command.Execute()
    
    If RecordSet.EOF = False Then
        Let ExecuteScalar = RecordSet.Fields(0).Value
    End If
  
    ClearParameter
    Sql.Clear
    
End Function

Public Function ExecuteNonQuery(ByVal CommandSql As String) As Boolean
    
    Dim RecordsAffected As Integer
    
    On Error GoTo Exception
            
    With Command
                
        .ActiveConnection = Connection
        .CommandType = adCmdText
        .CommandText = CommandSql
                
        Call .Execute(RecordsAffected)
            
    End With
            
    Let ExecuteNonQuery = RecordsAffected
            
    ClearParameter
    Sql.Clear
    
    Exit Function
    
Exception:
    Connection.RollbackTrans
    
End Function

Public Sub RollbackTrans()
    Connection.RollbackTrans
End Sub

Public Sub AddParameter(ByVal Name As String, Value As Variant, Optional DataType As DataTypeEnum = adEmpty, Optional ParameterDirection As ParameterDirectionEnum = adParamInput, Optional Size As Integer = 255)
    Command.Parameters.Append Command.CreateParameter(Name, DataType, ParameterDirection, Size, Value)
End Sub

Public Sub ClearParameter()
    While Command.Parameters.Count > 0
        Call Command.Parameters.Delete(0)
        DoEvents
    Wend
End Sub

Public Property Get Field(Name As String) As Variant
    Field = RecordSet.Fields(Name).Value
End Property

Private Sub Class_Terminate()
    Set Connection = Nothing
    Set RecordSet = Nothing
    Set Command = Nothing
    Set Sql = Nothing
End Sub
