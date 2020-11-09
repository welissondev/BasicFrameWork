VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccessTransactionConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Connection As ADODB.Connection
Public RecordSet As ADODB.RecordSet
Public Command As ADODB.Command
Public SqlBuilder As StringBuilder

Private Sub Class_Initialize()
    Set Connection = New ADODB.Connection
    Let Connection.CursorLocation = adUseClient
    Set Command = New ADODB.Command
    Set RecordSet = New ADODB.RecordSet
    Set SqlBuilder = New StringBuilder
End Sub

Public Function Begin() As ADODB.Connection
    
    With Connection
        If .State = False Then
           .Open AppProperty.ConnectionString
           .BeginTrans
        End If
    End With
    
    Set Begin = Connection
    
End Function

Public Sub Commit()
    Connection.CommitTrans
End Sub

Public Function ExecuteNonQuery(ByVal CommandSql As String) As Boolean
    
    Dim RecordsAffected As Integer
    
    On Error GoTo Exception
            
        With Command
            
            .ActiveConnection = Connection
            .CommandType = adCmdText
            .CommandText = CommandSql
            
            Call .Execute(RecordsAffected)
        
        End With
        
        ExecuteNonQuery = RecordsAffected
        
        ClearParameter
        SqlBuilder.Clear
    
    Exit Function
    
Exception:
        
    Connection.RollbackTrans
    
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
    SqlBuilder.Clear
    
End Function

Public Sub AddParameter(ByVal Name As String, Value As Variant, Optional DataType As DataTypeEnum = adEmpty, Optional ParameterDirection As ParameterDirectionEnum = adParamInput, Optional Size As Integer = 255)
    Command.Parameters.Append Command.CreateParameter(Name, DataType, ParameterDirection, Size, Value)
End Sub

Public Sub ClearParameter()
    While Command.Parameters.Count > 0
        Call Command.Parameters.Delete(0)
        DoEvents
    Wend
End Sub

Private Sub Class_Terminate()
   
    If Connection.State = True Then: Connection.Close
    If RecordSet.State = True Then: RecordSet.Close
    
    Set Connection = Nothing: Set RecordSet = Nothing
    Set Command = Nothing: Set SqlBuilder = Nothing
   
End Sub


