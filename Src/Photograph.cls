VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Photograph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Frabasic")
Option Explicit

Public FileString As String

Public Function GetFilePath() As String
    FileString = Application.GetOpenFilename("All Image files,*.jpg;*.jpg)")
    GetFilePath = FileString
End Function

Public Sub LoadFile(ByVal Image As MSForms.Image, Optional ByVal Path As String = Empty)
    Select Case Path
        Case Is = Empty
            If VerifyDirectoryFile(FileString) = True Then
                Image.Picture = LoadPicture(FileString)
            End If
        Case Is <> Empty
            If VerifyDirectoryFile(Path) = True Then
                Image.Picture = LoadPicture(Path)
            End If
    End Select
End Sub

Public Function CopyFile(ByVal Sourc As String, ByVal dest As String, Optional ByVal OverWiriteFile = True) As Boolean
    On Error GoTo Exception
    With New FileSystemObject
        Call .CopyFile(Sourc, dest, OverWiriteFile)
    End With
Success:
    CopyFile = True
    Exit Function
Exception:
    CopyFile = False
End Function

Public Function DeleteFile(ByVal Path As String) As Boolean
    On Error GoTo Exception
    Call Kill(Path)
Success:
    DeleteFile = True
    Exit Function
Exception:
    DeleteFile = False
End Function

Public Function VerifyDirectoryFile(ByVal Path As String) As Boolean
    With New FileSystemObject
        VerifyDirectoryFile = .FileExists(Path)
    End With
End Function

Public Function GetFileName(Path As String) As String
    With New FileSystemObject
        GetFileName = .GetFileName(Path)
    End With
End Function

