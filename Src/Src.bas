Attribute VB_Name = "Src"
'/*
'* *************************************************************************************
'* Site: www.diarioexcel.com.br
'* Contato: welisson@diarioexcel.com.br
'* Youtube: https://www.youtube.com/channel/UCSJAAxUzTj-qVVIKaqswQww?view_as=subscriber
'* *************************************************************************************
'*
'* Esse módulo possui propriedades e métodos para manipular os objetos do projeto
'* facilitando a exportação e importação dos componentes. Dessa forma, podemos versionar
'* nosso código font no GIT sem muitas complicações.
'*
'* Algumas dependências são obrigatórias para conseguir utilizar essa ferramenta. O primeiro
'* passo é acessar a guia de <Referências> do VBA e adicionar as seguintes bibliotecas: <Microsoft vba
'* extensibility 5.3 library> e <Microsoft Scripting Runtime>.
'*
'* É necessário habilitar o modo de segurança do excel em: <Arquivo\Opções\Central de confiabilidade
'* configurações da central de confiabilidade\Configuração do actviX\Modo de segurança>. Isso é
'* necessário, caso contrário não ira funcionar. É importante lembrar que essas configurações são
'* apenas para desenvolverdores, já mais devem ser aplicadas no abiente do cliente.
'*
'*/

Option Explicit
Option Private Module

'//Cria diretórios
Sub Init()
On Error GoTo ErrorFail

    Dim List As Object
    Dim RootFolder As String
    Dim Directory As Variant
    
    Let RootFolder = ThisWorkbook.Path
    Set List = CreateObject("System.Collections.ArrayList")
    
    With List
        .Add RootFolder & "\Src"
        .Add RootFolder & "\Def"
    End With
    
    With New FileSystemObject
        For Each Directory In List
            If Not .FolderExists(Directory) Then Call .CreateFolder(Directory): Debug.Print "Created directory -> "; Directory
        Next
    End With
    
    Set List = Nothing
    
ErrorExit:
    Debug.Print ""
    Exit Sub

ErrorFail:
    Debug.Print "   > Error: " & Err.Description
    
End Sub

'//Ignora componentes
Sub Ignore(ParamArray Components() As Variant)
On Error GoTo ErrorFail
    
    Dim RootFolder As String
    Dim File As TextStream
    Dim Name As Variant
    
    Let RootFolder = ThisWorkbook.Path & "\Def"
    
    With New FileSystemObject
        
        Set File = .OpenTextFile(RootFolder & "\Ignoreds.config", ForAppending, True)
        
        For Each Name In Components
             File.WriteLine Name
        Next
        
        File.Close
         
    End With

ErrorExit:
    Debug.Print ""
    Exit Sub

ErrorFail:
    Debug.Print "   > Error: " & Err.Description
    
End Sub

'//Exporta componentes
Sub Push()
On Error GoTo ErrorFail
    
    Dim Project As VBIDE.VBComponents
    Dim Component As VBIDE.VBComponent
    Dim RootFolder As String
    Dim List As Object
    
    Set Project = ThisWorkbook.VBProject.VBComponents
    Let RootFolder = ThisWorkbook.Path & "\Src\"
    Set List = GetIgnoredList()
    
    For Each Component In Project
        
        If Not List.Contains(Component.Name) Then
            
            Select Case ExtensionType(Component)
                Case Is = ".cls": Component.Export RootFolder & Component.Name & ".cls"
                Case Is = ".bas": Component.Export RootFolder & Component.Name & ".bas"
                Case Is = ".frm": Component.Export RootFolder & Component.Name & ".frm"
            End Select
        
        End If
        
        DoEvents
         
    Next

ErrorExit:
    Debug.Print ""
    Exit Sub
    
ErrorFail:
    Debug.Print "   > Error: " & Err.Description
    
End Sub

'\\Importa componentes
Sub Pull()
On Error GoTo ErrorFail
    
    Dim Project As VBIDE.VBComponents
    Dim Component As VBIDE.VBComponent
    Dim FileName As Variant
    Dim RootFolder, BaseName, Extension As String
    Dim FileList, IgnoredList As Object
    
    Set Project = ThisWorkbook.VBProject.VBComponents
    Set FileList = GetSrcFileList()
    Set IgnoredList = GetIgnoredList()
    
    '//Remove componentes atuais
    For Each Component In Project
        
        If Not IgnoredList.Contains(Component.Name) Then
        
            Select Case ExtensionType(Component)
                Case Is = ".cls":  Project.Remove Component
                Case Is = ".bas":  Project.Remove Component
                Case Is = ".frm":  Project.Remove Component
            End Select
            
        End If
        
    Next
    
    '//Importa componentes
    With New FileSystemObject
        
        Let RootFolder = ThisWorkbook.Path & "\Src\"
        
        For Each FileName In FileList
            
            BaseName = .GetBaseName(FileName)
            Extension = "." & .GetExtensionName(FileName)
            
            If Not IgnoredList.Contains(BaseName) Then
            
                Select Case Extension
                     Case Is = ".cls": Project.Import RootFolder & FileName
                     Case Is = ".bas": Project.Import RootFolder & FileName
                     Case Is = ".frm": Project.Import RootFolder & FileName
                End Select
            
            End If
            
            DoEvents
            
        Next
        
    End With
    
ErrorExit:
    Debug.Print ""
    Exit Sub

ErrorFail:
    Debug.Print "   > Error: " & Err.Description

End Sub


'//Atualiza diretório src
Sub Refresh()
On Error GoTo ErrorFail

    Dim SrcFileList, VbaFileList, IgnoredList As Object
    Dim FileName, BaseName, Extension As Variant
    Dim RootFolder As String

    Set SrcFileList = GetSrcFileList()
    Set VbaFileList = GetVbaFileList()
    Set IgnoredList = GetIgnoredList()
    
    Let RootFolder = ThisWorkbook.Path & "\Src\"
    
    With New FileSystemObject
        
        '//Remove todos os arquivos do src que não estão no projeto, ou
        '//que foram ignorados pelo usuário
        For Each FileName In SrcFileList
            
            If Not VbaFileList.Contains(FileName) Or IgnoredList.Contains(.GetBaseName(FileName)) = True Then
                
                Let BaseName = .GetBaseName(FileName)
                Let Extension = "." & .GetExtensionName(FileName)
                
                Select Case Extension
                    Case Is = ".cls"
                        .DeleteFile (RootFolder & FileName)
                    Case Is = ".bas"
                        .DeleteFile (RootFolder & FileName)
                    Case Is = ".frm"
                        .DeleteFile (RootFolder & FileName)
                        .DeleteFile (RootFolder & BaseName & ".frx")
                End Select
                
            End If
            
            DoEvents
            
        Next
        
    End With
    
ErrorExit:
    Exit Sub
    
ErrorFail:
    Debug.Print "   > Error: " & Err.Description
    
End Sub


'//Retorna <ArrayList> com nomes dos componentes do src
Function GetSrcFileList(Optional ViewImmediateWindows As Boolean = False) As Object
On Error GoTo ErrorFail
        
    Dim FileList, RootFolder As Object
    Dim FileName, Extesion As Variant

    Set FileList = CreateObject("System.Collections.ArrayList")
    
    With New FileSystemObject
        
        Set RootFolder = .GetFolder(ThisWorkbook.Path & "\Src\")
        
        For Each FileName In RootFolder.Files
            
            FileName = .GetFileName(FileName)
            
            Select Case .GetExtensionName(FileName)
                Case Is = "cls": FileList.Add FileName
                Case Is = "bas": FileList.Add FileName
                Case Is = "frm": FileList.Add FileName
            End Select
            
            DoEvents
            
        Next
        
    End With
    
    If ViewImmediateWindows = True Then
        Debug.Print ""
        Debug.Print "   Src Directory: " & RootFolder
        Debug.Print "   " & FileList.Count; " files found in the branch"
        Debug.Print ""
        For Each FileName In FileList
            Debug.Print "   -> "; FileName
            DoEvents
        Next
    End If
    
    Set GetSrcFileList = FileList
    
ErrorExit:
    Debug.Print ""
    Exit Function
    
ErrorFail:
    Debug.Print "   > Error: " & Err.Description

End Function

'//Retorna <ArrayList> com nomes dos componentes do projeto
Function GetVbaFileList(Optional ViewImmediateWindows As Boolean = False) As Object
On Error GoTo ErrorFail

    Dim Component As VBComponent
    Dim Project As VBComponents
    Dim FileList As Object
    Dim FileName As Variant
    
    Set Project = ThisWorkbook.VBProject.VBComponents
    Set FileList = CreateObject("System.Collections.ArrayList")
    
    For Each Component In Project
        
        Select Case ExtensionType(Component)
            Case Is = ".cls": FileList.Add Component.Name & ".cls"
            Case Is = ".bas": FileList.Add Component.Name & ".bas"
            Case Is = ".frm": FileList.Add Component.Name & ".frm"
        End Select
        
        DoEvents
        
    Next
    
    If ViewImmediateWindows = True Then
        Debug.Print ""
        Debug.Print "   ThisWorkbook: "; ThisWorkbook.VBProject.Name
        Debug.Print "   " & FileList.Count; " files found in the project"
        Debug.Print ""
        For Each FileName In FileList
            Debug.Print "   -> "; FileName
            DoEvents
        Next
    End If

    Set GetVbaFileList = FileList

ErrorExit:
    Debug.Print ""
    Exit Function
    
ErrorFail:
    Debug.Print "   > Error: " & Err.Description

End Function

Function GetIgnoredList(Optional ViewImmediateWindows As Boolean = False) As Object
On Error GoTo ErrorFail

    Dim Path As String
    Dim Stream As TextStream
    Dim FileName As Variant
    Dim ArrayList As Object
    
    Set ArrayList = CreateObject("System.Collections.ArrayList")
    Let Path = ThisWorkbook.Path & "\Def\Ignoreds.config"
    
    With New FileSystemObject
        
        Set Stream = .OpenTextFile(Path, ForReading, True)
        
        While Not Stream.AtEndOfLine
            
            Let FileName = Stream.ReadLine
            ArrayList.Add FileName
            
        Wend
        
    End With
    
    If ViewImmediateWindows = True Then
        Debug.Print ""
        Debug.Print "   Ignored.Config: "; Path
        Debug.Print "   " & ArrayList.Count; " ignored files"
        Debug.Print ""
        For Each FileName In ArrayList
            Debug.Print "   -> "; FileName
            DoEvents
        Next
    End If
    
    Set GetIgnoredList = ArrayList

ErrorExit:
    Debug.Print ""
    Exit Function

ErrorFail:
    Debug.Print "   > Error: " & Err.Description

End Function

'//Obtem o tipo de extensão
Private Property Get ExtensionType(Component As VBIDE.VBComponent) As String
    Select Case Component.Type
        Case Is = vbext_ct_ClassModule
            ExtensionType = ".cls"
        Case Is = vbext_ct_StdModule
            ExtensionType = ".bas"
        Case Is = vbext_ct_MSForm
            ExtensionType = ".frm"
    End Select
End Property

