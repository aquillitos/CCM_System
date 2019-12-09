Attribute VB_Name = "03_File_Control"
Option Compare Database:    Option Explicit

Public Function f_Folder_Exist(folderPath As String) As Boolean
    On Error GoTo ERR_EXIT
    f_Folder_Exist = False
    If Dir(folderPath, vbDirectory) = "" Then
            f_Folder_Exist = False
    Else
            f_Folder_Exist = True
    End If
    Exit Function
    
ERR_EXIT:
    f_Folder_Exist = False
    Select Case ERR
            Case 52
                    ANS = MsgBox("You don't seem to have a permission to the folder.", vbCritical + vbOKOnly)
            Case Else
                    ANS = MsgBox("Error: Please ask the System Administrator.", vbCritical + vbOKOnly)
    End Select
End Function

Public Function f_Create_Folder(FolderName As String) As String
        On Error GoTo ERR_EXIT
        Dim objFileSys As Object
        Dim strCreateFolder As String
        Dim sFolder As String

        Call s_DeskTopPath
        sFolder = DeskTopPath & "\" & FolderName
            
        Set objFileSys = CreateObject("Scripting.fileSystemobject")
        If objFileSys.FolderExists(sFolder) = False Then
                strCreateFolder = objFileSys.BuildPath(DeskTopPath, FolderName)
                objFileSys.CreateFolder strCreateFolder
        End If
        f_Create_Folder = sFolder
        Set objFileSys = Nothing:  strCreateFolder = ""
        Exit Function
        
ERR_EXIT:
        f_Create_Folder = ""
        Set objFileSys = Nothing:  strCreateFolder = ""
End Function

Public Function f_Create_AttachmentFolder(pFolderName, cFolderName) As Boolean
        f_Create_AttachmentFolder = False
        Dim objFileSys As Object
        Dim strCreateFolder As String
        Dim sFolder As String
        
        On Error GoTo ERR_EXIT
        If Dir(pFolderName & "\" & cFolderName, vbDirectory) = "" Then
                Set objFileSys = CreateObject("Scripting.fileSystemobject")
                If objFileSys.FolderExists(sFolder) = False Then
                        strCreateFolder = objFileSys.BuildPath(pFolderName, cFolderName)
                        objFileSys.CreateFolder strCreateFolder
                End If
                f_Create_AttachmentFolder = True
                Set objFileSys = Nothing:  strCreateFolder = ""
        End If
        Exit Function
        
ERR_EXIT:
    f_Create_AttachmentFolder = False
    Select Case ERR
            Case 52
                    ANS = MsgBox("You don't seem to have a permission to the folder.", vbCritical + vbOKOnly)
            Case Else
                    ANS = MsgBox("Error: Please ask the System Administrator.", vbCritical + vbOKOnly)
    End Select
End Function

Public Function f_DeleteFiles(fileName As String) As Boolean
        On Error GoTo ERR_EXIT
        f_DeleteFiles = False
        
        Call f_Folder_Exist(fileName)
        If Len(Dir(fileName)) > 0 Then
                Kill fileName
                f_DeleteFiles = True
                Exit Function
        End If
    
ERR_EXIT:
        Select Case ERR
                Case 52
                        ANS = MsgBox("You don't seem to have a permission to the folder.", vbCritical + vbOKOnly)
                        f_DeleteFiles = False
                Case Else
                        'ANS = MsgBox("Error: Please ask the System Administrator.", vbCritical + vbOKOnly)
                        f_DeleteFiles = True
        End Select
End Function

Public Function f_DeleteDirectory(directory As String)
        On Error Resume Next
        
        'Directory の中身チェック
        Dim insideDir As String
        Dim cCount As Integer
        cCount = 0
        insideDir = Dir(directory & "\*.*", vbNormal)
        Do Until insideDir = ""
                insideDir = Dir()
                cCount = cCount + 1
        Loop
        
        'Directoryの中身が無かったら、Directoryを消去する。
        If cCount = 0 Then
                Dim Fso As FileSystemObject 'Reference「Microsoft Scripting Runtime」 が必要
                Set Fso = New FileSystemObject
                Call Fso.DeleteFolder(directory, True)
                Set Fso = Nothing
        End If
        insideDir = ""
        cCount = 0
        
        On Error GoTo 0
End Function

Public Function f_File_PickUp(tTitle As String, iName As String) As Boolean
        f_File_PickUp = False
        
        Dim MFile As MsoFileDialogType ' Requires Reference : Microsoft Office 15.0 Object Library
        Dim FDialog As FileDialog
        
        Call s_DeskTopPath
        MFile = msoFileDialogFilePicker
        Set FDialog = FileDialog(MFile)
        With FDialog
                .AllowMultiSelect = False
                .Title = tTitle
                .ButtonName = "Import"
                .InitialFileName = iName
        End With
            
        FDialog.Show
        If FDialog.SelectedItems.count = 1 Then
                f_File_PickUp = True
                filePath = FDialog.SelectedItems(1)
                'Debug.Print filePath
        End If
End Function

Public Function f_Import_Data(ImportName As String, filePath As String) As Boolean
        f_Import_Data = False
        Call f_Delete_LocalTable(ImportName)
        On Error GoTo e
        DoCmd.TransferSpreadsheet acImport, , ImportName, filePath, True
        f_Import_Data = True
        Exit Function
e:
        f_Import_Data = False
End Function
