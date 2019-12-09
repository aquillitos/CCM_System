Attribute VB_Name = "01_System_Configuration"
Option Compare Database:    Option Explicit

Public Function s_Get_System_Config()
        Dim stSQL1 As String
        Dim rs1 As DAO.Recordset
        
        stSQL1 = "SELECT * FROM [CCM_MST_System_Config];"
        Set rs1 = db.OpenRecordset(stSQL1, dbReadOnly)
        Do Until rs1.EOF
                Select Case rs1![ID]
                        Case Is = 1 'DB_SQL_Connect
                                ConSQL = rs1![Value]
                        Case Is = 2 'DB_ADO_Connect
                                ConADO = rs1![Value]
                        Case Is = 3 'Attach_Directory
                                attachDir = rs1![Value]
                        Case Else
                End Select
                rs1.MoveNext
        Loop
        rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
End Function

Public Function s_System_Server_Connect() As Boolean  'SQL server
        s_System_Server_Connect = False
        On Error GoTo ErrSub
                Set ConSys = New ADODB.Connection
                ConSys.ConnectionString = ConADO
                ConSys.Open
        On Error GoTo 0
        s_System_Server_Connect = True
        Exit Function
ErrSub:
        Call s_Server_Connect_Fail("SQL Server Connect Failed.")
        s_System_Server_Connect = False
End Function

Public Function s_Local_Server_Connect() As Boolean
        s_Local_Server_Connect = False
        On Error GoTo ErrSub
            Set ConLo = CurrentProject.Connection
        On Error GoTo 0
        s_Local_Server_Connect = True
        Exit Function
ErrSub:
        Call s_Server_Connect_Fail("Local Server Connect Failed.")
        s_Local_Server_Connect = False
End Function

Public Function s_Server_Connect_Fail(C_msgTitle As String)
    ANS = MsgBox("Server Connection Failed." & vbCrLf & "Please check the Server Status and the Connection String then try again", vbCritical + vbOKOnly, C_msgTitle)
End Function

Public Function s_Login_ID() As Boolean
        s_Login_ID = False
        Dim stSQL1 As String, stSQL2 As String
        Dim rs1 As ADODB.Recordset, rs2 As ADODB.Recordset
        
        Set rs1 = New ADODB.Recordset
        stSQL1 = "SELECT * FROM CCM.[" & MST01 & "] WHERE [LANID] = '" & CreateObject("WScript.Network").userName & "';"
        rs1.Open stSQL1, ConSys, adOpenForwardOnly, adLockReadOnly
        If Not rs1.EOF Then
                userID = rs1![LANID]
                userName = rs1![User_Name]
                userMail = rs1![eMail]
                
                stSQL2 = "SELECT * FROM CCM.[" & MST02 & "] WHERE ID = " & rs1![authentication_ID] & ";"
                Set rs2 = New ADODB.Recordset
                rs2.Open stSQL2, ConSys, adOpenForwardOnly, adLockReadOnly
                If Not rs2.EOF = True Then
                        authName = rs2![Name]
                        authMaster = rs2![f_Master]
                        authProperty = rs2![f_Property]
                        authCost = rs2![f_Cost]
                        authOther = rs2![f_Other]
                        authCase1 = rs2![f_Case1]
                        authLevel = rs2![authentication_level]
                End If
                rs2.Close: Set rs2 = Nothing:   stSQL2 = ""
                s_Login_ID = True
        Else
                ANS = MsgBox("You do not have any access to the system." & vbCrLf & "Please ask the administrator for your access." & vbCrLf & "The application is aborted.", vbCritical + vbOKOnly)
                DoCmd.Close acForm, FRM01, acSaveNo
                DoCmd.Quit acQuitPrompt
                s_Login_ID = False
        End If
        rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
End Function

Public Function s_DeskTopPath() As String
        DeskTopPath = ""
        Dim WSH As Object
        Set WSH = CreateObject("Wscript.Shell")
        DeskTopPath = WSH.SpecialFolders("Desktop")
        Set WSH = Nothing
End Function

Public Function s_System_Config_Update()
        If f_Copy_TableSL(MST00, "TMP_MST") = True Then
                If f_Delete_LocalTable(MST00) = True Then
                        DoCmd.Rename MST00, acTable, "TMP_MST"
                End If
        End If
End Function

