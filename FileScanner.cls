VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/*
'* *************************************************************************************
'* Site: www.diarioexcel.com.br
'* Contato: welisson@diarioexcel.com.br
'* Youtube: https://www.youtube.com/channel/UCSJAAxUzTj-qVVIKaqswQww?view_as=subscriber
'* *************************************************************************************
'*
'* Essa classe pode ser utilizada para retornar uma coleção de arquivos
'* de um diretório especificado <UserFolderPath>, incluindo também os arquivos
'* contidos nos sub-diretórios da pasta root.
'*
'* Você pode passar como parâmetro o <ExtensionFilter> caso deseje filtrar
'* arquivos específicos do diretório... Ex: Array(".txt",".log",".exe"),
'* se nenhum parâmetro for passado a classe retornará tudo.
'*
'* É possível definir o <IncludeSubDirectory> para que o sistema
'* retorne apenas os arquivos do diretório raiz, por padrão esse
'* parâmetro é opcional <OptIncludeSubDirectory>
'*
'* A classe disponibiliza opções de anexeação com <ScrReturnOptions>,
'* dessa forma você pode definir como deseja retornar um arquivo. Por exemplo, se você
'* precisa do <name.extension> passe o parâmetro como <OptFileName>.
'*
'*/

'//Opções de retorno do arquivo
Public Enum ScrReturnOptions
    OptAbsolutePathName = 0
    OptFileName = 1
    OptBaseName = 2
    OptExtensionName = 3
End Enum

'//Opção de inclusão do subdiretório
Public Enum ScrDirectoryOptions
    OptIncludeSubDirectory = True
    OptNotIncludeSubDirectory = False
End Enum


Private FileSystemObject As Object
Private Files As Object


Private Sub Class_Initialize()
    Set Files = CreateObject("System.Collections.ArrayList")
    Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
End Sub

Public Function ScanFiles(ByVal UserFolderPath As String, Optional ByVal ExtensionFilter As Variant = Null, _
Optional ByVal IncludeSubDirectory As ScrDirectoryOptions = OptIncludeSubDirectory, Optional ByVal SetReturnOption As _
ScrReturnOptions = OptAbsolutePathName) As Object
    
    Dim File As Object
    Dim FileExtension As String
    Dim RootFolder As Object
    Dim RootSubFolder As Object
    
    '//Definido como statico por causa da chamada recursiva
    Static ScannedRootFolder As Boolean
        
    Set RootFolder = FileSystemObject.GetFolder(UserFolderPath)
    
    If Not ScannedRootFolder Then
        
        For Each File In RootFolder.Files
            
            Let FileExtension = "." & FileSystemObject.GetExtensionName(File)
            
            '//Se </ExtensionFilter> for null anexa todos os arquivos
            Select Case IsNull(ExtensionFilter)
                Case Is = True
                    Call Append(File, SetReturnOption)
                    
                Case Is = False
                    If CheckExtensionFile(FileExtension, ExtensionFilter) = True Then
                        Call Append(File, SetReturnOption)
                    End If
            End Select
          
        Next
                              
        '//Finaliza procedimento se não foi incluso sub-diretórios
        If Not IncludeSubDirectory Then
            Set ScanFiles = Files: Exit Function
        End If
        
        '//Executa chamada recursiva se pasta raiz contiver sub-pastas
        If RootFolder.SubFolders.Count > 0 Then
            Let ScannedRootFolder = True: Call ScanFiles(UserFolderPath, ExtensionFilter, _
            IncludeSubDirectory, SetReturnOption)
        End If
                     
    Else
             
        For Each RootSubFolder In RootFolder.SubFolders
            
            For Each File In RootSubFolder.Files
                                
                Let FileExtension = "." & FileSystemObject.GetExtensionName(File)
                
                '//Se </ExtensionFilter> for null anexa todos os arquivos das sub-pastas
                Select Case IsNull(ExtensionFilter)
                    Case Is = True
                        Call Append(File, SetReturnOption)
                        
                    Case Is = False
                        If CheckExtensionFile(FileExtension, ExtensionFilter) = True Then
                            Call Append(File, SetReturnOption)
                        End If
                End Select
                
                DoEvents
            Next
            
            '//Executa chamada recursiva se sub-pastas contiverem sub-pastas
            If RootSubFolder.SubFolders.Count > 0 Then
                Call ScanFiles(UserFolderPath & "\" & RootSubFolder.Name, ExtensionFilter, _
                IncludeSubDirectory, SetReturnOption)
            End If
            
            DoEvents
        Next
    
    End If
    
    Set ScanFiles = Files
    
End Function

Private Function CheckExtensionFile(ByVal ExtensionFile As String, ByVal ExtensionFilter As Variant) As Boolean
    
    Dim Index As Integer
    
    For Index = LBound(ExtensionFilter) To UBound(ExtensionFilter)
        
        If LCase(ExtensionFile) Like "*" & LCase(ExtensionFilter(Index)) & "*" Then
                CheckExtensionFile = True
            Exit For
        End If
        
        DoEvents
    Next
    
End Function


Private Sub Append(ByVal File As Object, ByVal SetReturnOption As ScrReturnOptions)

    Select Case SetReturnOption
        Case Is = 0
            Files.Add FileSystemObject.GetAbsolutePathName(File)
        Case Is = 1
            Files.Add FileSystemObject.GetFileName(File)
        Case Is = 2
            Files.Add FileSystemObject.GetBaseName(File)
        Case Is = 3
            Files.Add FileSystemObject.GetExtensionName(File)
    End Select
     
End Sub


Private Sub Class_Terminate()
    Set Files = Nothing
    Set FileSystemObject = Nothing
End Sub

