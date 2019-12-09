Attribute VB_Name = "14_Attachment"
Option Compare Database: Option Explicit

Public Function f_List_Attachment(cID As Long) As Boolean
        f_List_Attachment = False
        Dim stSQL1 As String
        With Forms(FRM03)
                stSQL1 = "SELECT * FROM " & ConSQL & ".[CCM." & CCMAttach & "] WHERE [CCM_ID] = " & cID & " ORDER BY [fileIndex];"
                .Text301.RowSource = stSQL1
        End With
End Function

Public Function f_Attachment_Check(contractNumber As String)
        'DBにあってFolderに無かった場合、DBを消す
        Dim stSQL1 As String, stSQL2 As String
        Dim rs1 As ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        
        stSQL1 = "SELECT * FROM CCM.[" & CCMAttach & "] WHERE [CCM_ID] = " & contractID & ";"
        rs1.Open stSQL1, ConSys, adOpenDynamic, adLockOptimistic
        Do Until rs1.EOF
                If Len(Dir(rs1![fileDirectory])) = 0 Then
                        stSQL2 = "DELETE FROM " & ConSQL & ".[CCM." & CCMAttach & "] WHERE [ID] = " & rs1![ID] & ";"
                        Call f_RunQuery(stSQL2):    stSQL2 = ""
                End If
                rs1.MoveNext
        Loop
        rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
        
        'DBになくてFolderにあった場合、DBに追加する
            '===> 削除： Folder内の命名規則がDBに合わない場合、DBに追加できない。
            '===> Folderへの格納は、Systemからのみ可能とする。
End Function

Public Function f_Attachment_Upload(contractID As Long, originalFilePath As String) As Boolean
        On Error GoTo ER
        f_Attachment_Upload = False
        
        If f_Create_AttachmentFolder(attachDir, contractNumber) = True Then
                Dim originalFileName As String, saveFileName As String, saveFilePath As String
                Dim stSQL1 As String, stSQL2 As String
                Dim rs1 As ADODB.Recordset, rs2 As ADODB.Recordset
                Set rs1 = New ADODB.Recordset
                Set rs2 = New ADODB.Recordset
                        originalFilePath = originalFilePath
                        originalFileName = Right(originalFilePath, Len(originalFilePath) - InStrRev(originalFilePath, "\", -1, vbTextCompare))
                        Debug.Print originalFileName & " - " & originalFilePath
                
                        'Check Attachment File Data and get the fileIndex ================================================
                        stSQL1 = "SELECT * FROM CCM.[" & CCMAttach & "] WHERE [CCM_ID] = " & contractID & " AND [fileName] = '" & originalFileName & "';"
                        rs1.Open stSQL1, ConSys, adOpenDynamic, adLockOptimistic
                                If Not rs1.EOF Then 'If the same file exists. --------------------------------------------------------
                                        fileindex = rs1![fileindex]
                                        saveFilePath = attachDir & "\" & contractNumber & "\" & contractNumber & "-" & fileindex & "-" & originalFileName
                                            rs1![fileName] = originalFileName
                                            rs1![fileDirectory] = saveFilePath
                                            rs1![Update] = Now()
                                            rs1![update_by] = "TEST"
                                        rs1.Update
                                        ANS = MsgBox("The Attachment has replaced with the new one.", vbInformation + vbOKOnly)
                                Else    'If the same file doesn't exist. -----------------------------------------------------------------
                                        stSQL2 = "SELECT MAX([fileIndex]) AS maxIndx FROM CCM.[" & CCMAttach & "] WHERE [number] = '" & contractNumber & "' ;"
                                        rs2.Open stSQL2, ConSys, adOpenDynamic, adLockReadOnly
                                                If rs2.EOF Then 'If this is the first attachment, the [fileIndex] is 1. -------------------
                                                        fileindex = 1
                                                Else    'If not the first attachment, get the Max [fileIndex] --------------------------------
                                                        fileindex = Nz(rs2![maxIndx], 0) + 1
                                                End If
                                        rs2.Close:  Set rs2 = Nothing:  stSQL2 = ""
                                        saveFilePath = attachDir & "\" & contractNumber & "\" & contractNumber & "-" & fileindex & "-" & originalFileName
                                        rs1.AddNew
                                            rs1![CCM_ID] = contractID
                                            rs1![number] = contractNumber
                                            rs1![fileindex] = fileindex
                                            rs1![fileName] = originalFileName
                                            rs1![fileDirectory] = saveFilePath
                                            rs1![Update] = Now()
                                            rs1![update_by] = "TEST"
                                        rs1.Update
                                End If
                        rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
        
                        'Delete the existing file in the directory ===============================================
                        Call f_DeleteFiles(saveFilePath)
                        
                        'Copy to the directory ===========================================================
                        FileCopy originalFilePath, saveFilePath
                        
                        f_Attachment_Upload = True
        Else
ER:
                        f_Attachment_Upload = False
        End If
End Function

Public Function f_Attachment_Download(contractNumber As String) As Boolean
        Dim tFolder As String, oFolder As String
        Dim Fso As Object
        oFolder = attachDir & "\" & contractNumber
        tFolder = DeskTopPath & "\" & contractNumber
        
        If f_Folder_Exist(oFolder) = True Then
                'On Error GoTo ER
                Set Fso = CreateObject("Scripting.FileSystemObject")
                Fso.CopyFolder oFolder, tFolder
                ANS = MsgBox("Downloaded onto Desktop.", vbInformation + vbOKOnly)
        End If
ER:
        Set Fso = Nothing
        oFolder = "":   tFolder = ""
End Function

Public Function f_Attachment_Delete(attachID As Integer) As Boolean
        Debug.Print attachID
        Dim stSQL1 As String, stSQL2 As String
        Dim rs1 As ADODB.Recordset, rs2 As ADODB.Recordset
        Dim pDirectory As String
        Set rs1 = New ADODB.Recordset
        Set rs2 = New ADODB.Recordset
        
        stSQL1 = "SELECT * FROM CCM.[" & CCMAttach & "] WHERE [ID] = " & attachID & ";"
        rs1.Open stSQL1, ConSys, adOpenDynamic, adLockOptimistic
        If Not rs1.EOF Then
                'Delete From Directory ====================================================
                If f_DeleteFiles(rs1![fileDirectory]) = True Then
                
                        'Delete the Whole Contract Directory if no file ================================
                        pDirectory = Left(rs1![fileDirectory], InStrRev(rs1![fileDirectory], "\", , vbTextCompare) - 1)
                        Call f_DeleteDirectory(pDirectory)
                        
                        'Delete From Data Table ===============================================
                        stSQL2 = "DELETE FROM CCM.[" & CCMAttach & "] WHERE [ID] = " & attachID & ";"
                        rs2.Open stSQL2, ConSys, adOpenDynamic, adLockOptimistic
                        Set rs2 = Nothing:  stSQL2 = ""
                Else
                        'On Error (No Permission, No folder etc =====================================
                        Exit Function
                End If
        End If
        rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
End Function

