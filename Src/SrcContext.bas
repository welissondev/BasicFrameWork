Attribute VB_Name = "SrcContext"
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
'* apenas para desenvolverdores, já mais devem ser aplicadas no abiente de produção.
'*
'*/

Option Explicit
Option Private Module

Enum CmdImmediate
    PrintLS = 1
End Enum

'//Cria diretórios
Sub Init()
On Error GoTo ErrorFail

    Dim NameList As Object
    Dim RootFolder As String
    Dim Directory As Variant
    
    Let RootFolder = ThisWorkbook.Path
    Set NameList = CreateObject("System.Collections.ArrayList")
    
    With NameList
        .Add RootFolder & "\Src"
        .Add RootFolder & "\Def"
    End With
    
    With New FileSystemObject
    
        For Each Directory In NameList
            If Not .FolderExists(Directory) Then
                .CreateFolder (Directory): Debug.Print "Created directory -> "; Directory
            End If
        Next
    
    End With
    
    Set NameList = Nothing
    
ErrorExit:
    Debug.Print ""
    Exit Sub

ErrorFail:
    Debug.Print "   > Error: " & Err.Description
    
End Sub

'//Ignora componentes
Sub Ignore(ParamArray ComponentNames() As Variant)
On Error GoTo ErrorFail
    
    Dim RootFolder As String
    Dim Stream As TextStream
    Dim Arg As Variant
    
    Let RootFolder = ThisWorkbook.Path & "\Def"
    
    With New FileSystemObject
        
        Set Stream = .OpenTextFile(RootFolder & "\Ignoreds.config", ForAppending, True)
        
        For Each Arg In ComponentNames
             Stream.WriteLine Arg
        Next
        
        Stream.Close
         
    End With

ErrorExit:
    Debug.Print ""
    Exit Sub

ErrorFail:
    Debug.Print "   > Error: " & Err.Description
    
End Sub

'//Empurra componentes
Sub Push()
On Error GoTo ErrorFail
    
    Dim Project As VBIDE.VBComponents
    Dim Component As VBIDE.VBComponent
    Dim RootFolder As String
    Dim NameList As Object
    
    Set Project = ThisWorkbook.VBProject.VBComponents
    Let RootFolder = ThisWorkbook.Path & "\Src\"
    Set NameList = GetIgnoreds()
    
    For Each Component In Project
        
        If Not NameList.Contains(Component.Name) Then
            
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

'\\Puxa componentes
Sub Pull()
On Error GoTo ErrorFail
    
    Dim Project As VBIDE.VBComponents
    Dim Component As VBIDE.VBComponent
    Dim FileName As Variant
    Dim RootFolder, BaseName, Extension As String
    Dim SrcList, IgnoredList As Object
    
    Let RootFolder = ThisWorkbook.Path & "\Src\"
    Set Project = ThisWorkbook.VBProject.VBComponents
    Set IgnoredList = GetIgnoreds()
    Set SrcList = GetSrcFiles()

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
        
        For Each FileName In SrcList
            
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
Sub Rebase()
On Error GoTo ErrorFail

    Dim SrcList, VbaList, IgnoredList As Object
    Dim FileName, BaseName, Extension As Variant
    Dim RootFolder As String

    Set SrcList = GetSrcFiles()
    Set VbaList = GetVbaFiles()
    Set IgnoredList = GetIgnoreds()
    
    Let RootFolder = ThisWorkbook.Path & "\Src\"
    
    '//Remove todos os arquivos do src que não estão no projeto, ou
    '//que foram ignorados pelo usuário
    With New FileSystemObject
        
        For Each FileName In SrcList
            
            Let BaseName = .GetBaseName(FileName)
            Let Extension = "." & .GetExtensionName(FileName)
            
            If Not VbaList.Contains(FileName) Or IgnoredList.Contains(BaseName) = True Then
                
                Select Case Extension
                    Case Is = ".cls"
                        .DeleteFile (RootFolder & FileName)
                    Case Is = ".bas"
                        .DeleteFile (RootFolder & FileName)
                    Case Is = ".frm"
                        .DeleteFile (RootFolder & FileName): .DeleteFile (RootFolder & BaseName & ".frx")
                End Select
                
            End If
            
            DoEvents
            
        Next
        
    End With
    
ErrorExit:
    Debug.Print
    Exit Sub
    
ErrorFail:
    Debug.Print "   > Error: " & Err.Description
    
End Sub

'//Retorna <ArrayList> com nomes dos componentes do src
Function GetSrcFiles(Optional ImmeView As CmdImmediate = 0) As Object
On Error GoTo ErrorFail

    Dim NameList, RootFolder As Object
    Dim FileName, Extesion As Variant

    Set NameList = CreateObject("System.Collections.ArrayList")
    
    With New FileSystemObject
            
        Set RootFolder = .GetFolder(ThisWorkbook.Path & "\Src\")
            
        For Each FileName In RootFolder.Files
            
            FileName = .GetFileName(FileName)
            
            Select Case .GetExtensionName(FileName)
                Case Is = "cls": NameList.Add FileName
                Case Is = "bas": NameList.Add FileName
                Case Is = "frm": NameList.Add FileName
            End Select
            
            DoEvents
            
        Next
        
    End With
    
    Call NameList.Sort
    
    If ImmeView = 1 Then
        Debug.Print ""
        Debug.Print "   Src Directory: " & ThisWorkbook.Path & "\Src"
        Debug.Print "   " & NameList.Count; " files found in the branch"
        Debug.Print ""
        For Each FileName In NameList
            Debug.Print "   -> "; FileName
            DoEvents
        Next
        Debug.Print ""
    End If
       
    Set GetSrcFiles = NameList
    
ErrorExit:
    Exit Function
    
ErrorFail:
    Debug.Print "   > Error: " & Err.Description

End Function

'//Retorna <ArrayList> com nomes dos componentes do projeto
Function GetVbaFiles(Optional ImmeView As CmdImmediate = 0) As Object
On Error GoTo ErrorFail
    
    Dim Project As VBComponents
    Dim Component As VBComponent
    Dim NameList As Object
    Dim FileName As Variant
    
    Set Project = ThisWorkbook.VBProject.VBComponents
    Set NameList = CreateObject("System.Collections.ArrayList")
    
    For Each Component In Project
        
        Select Case ExtensionType(Component)
            Case Is = ".cls": NameList.Add Component.Name & ".cls"
            Case Is = ".bas": NameList.Add Component.Name & ".bas"
            Case Is = ".frm": NameList.Add Component.Name & ".frm"
        End Select
        
        DoEvents
        
    Next
    
    Call NameList.Sort
    
    If ImmeView = 1 Then
        Debug.Print ""
        Debug.Print "   ThisWorkbook: "; ThisWorkbook.VBProject.Name & " (" & ThisWorkbook.Name & ")"
        Debug.Print "   " & NameList.Count; " files found in the project"
        Debug.Print ""
        For Each FileName In NameList
            Debug.Print "   -> "; FileName
            DoEvents
        Next
        Debug.Print ""
    End If

    Set GetVbaFiles = NameList

ErrorExit:
    Exit Function
    
ErrorFail:
    Debug.Print "   > Error: " & Err.Description

End Function

'//Retorna uma <ArrayList> com os nomes dos componentes ignorados
Function GetIgnoreds(Optional ImmeView As CmdImmediate = 0) As Object
On Error GoTo ErrorFail

    Dim Path As String
    Dim Stream As TextStream
    Dim FileName As Variant
    Dim IgnoredList As Object
    
    Set IgnoredList = CreateObject("System.Collections.ArrayList")
    Let Path = ThisWorkbook.Path & "\Def\Ignoreds.config"
    
    With New FileSystemObject
        
        Set Stream = .OpenTextFile(Path, ForReading, True)
        
        While Not Stream.AtEndOfLine
            
            Let FileName = Stream.ReadLine
            IgnoredList.Add FileName
            
        Wend
        
    End With
    
    Call IgnoredList.Sort
    
    If ImmeView = 1 Then
        Debug.Print ""
        Debug.Print "   Ignored.Config: "; Path
        Debug.Print "   " & IgnoredList.Count; " ignored files"
        Debug.Print ""
        For Each FileName In IgnoredList
            Debug.Print "   -> "; FileName
            DoEvents
        Next
        Debug.Print ""
    End If
    
    Set GetIgnoreds = IgnoredList

ErrorExit:
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
