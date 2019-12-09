Attribute VB_Name = "51_Download"
Option Compare Database:    Option Explicit

Public Function f_Download(cCount As Integer) As Boolean
        On Error GoTo ERR
        DoCmd.Hourglass True
        f_Download = False
        Dim stSQL1 As String, stSQL2 As String, stSQL3 As String, stSQL4 As String, stSQL5 As String
        Dim rs1 As DAO.Recordset
        Dim fileName As String
        
        'Delete Temporary Tables =====================================================================================
        Call f_Delete_LocalTable(TMP21)
        Call f_Delete_LocalTable(TMP22)
        Call f_Delete_LocalTable(TMP23)
        Call f_Delete_LocalTable(TMP24)
        
        'Set Export Folder ==========================================================================================
        stSQL1 = "SELECT [value] FROM " & MST00 & " WHERE [ID] = 4 ;"
        Set rs1 = db.OpenRecordset(stSQL1, dbReadOnly)
        If Not rs1.EOF Then
                fileName = DeskTopPath & "\" & rs1![Value] & "_" & Format(Now(), "yyyymmddhhmm")
        End If
        rs1.Close: Set rs1 = Nothing:   stSQL1 = ""
        
        'Data Table ===============================================================================================
        stSQL2 = "SELECT * INTO [" & TMP21 & "] FROM " & ConSQL & ".[CCM." & CCMDATA & "] " & f_Search_Query(cCount)
        ConLo.Execute stSQL2:   stSQL2 = ""
        
        'Price Table ==============================================================================================
        'stSQL3 = "SELECT [Q1].* INTO [" & TMP22 & "] FROM " & ConSQL & ".[CCM." & CCMCOST1 & "] AS [Q1] INNER JOIN " & _
                    "(SELECT [CCM_ID], Max([term_end]) AS [end_Max] " & _
                    "FROM " & ConSQL & ".[CCM." & CCMCOST1 & "] GROUP BY [CCM_ID] ) AS [Q2] " & _
                    "ON [Q1].[term_end] = [Q2].[end_Max] AND [Q1].[CCM_ID] = [Q2].[CCM_ID];"
        
        stSQL3 = "SELECT [Q1].* INTO [" & TMP22 & "] FROM " & ConSQL & ".[CCM." & CCMCOST1 & "] AS [Q1] INNER JOIN " & _
                    "(SELECT [CCM_ID], Max([ID]) AS [ID_Max] " & _
                    "FROM " & ConSQL & ".[CCM." & CCMCOST1 & "] GROUP BY [CCM_ID] ) AS [Q2] " & _
                    "ON [Q1].[ID] = [Q2].[ID_Max] AND [Q1].[CCM_ID] = [Q2].[CCM_ID];"
        ConLo.Execute stSQL3:   stSQL3 = ""
        
        'BOID Table ====== Added 2019/11/22 ==========================================================================
        stSQL4 = "SELECT [Q1].* INTO [" & TMP23 & "] FROM " & ConSQL & ".[CCM." & CCMBOID & "] AS [Q1] INNER JOIN " & _
                        "(SELECT MAX([year]) AS [Year_Max] FROM " & ConSQL & ".[CCM." & CCMBOID & "]) AS [Q2] " & _
                        "ON [Q1].[year] = [Q2].[Year_Max] ;"
        ConLo.Execute stSQL4:   stSQL4 = ""
        
        'Create Joined Table ===== Added [Monthly Amount] [Annual Amount], Deleted [Budget ID], [Budget Name] 2019/11/22 ================
        If f_Exist_LocalTable(TMP21) = True And f_Exist_LocalTable(TMP22) = True Then
                stSQL5 = "SELECT [Q1].*, " & _
                                "[Q2].[terms], [Q2].[currency], [Q2].[payment_schedule], " & _
                                "[Q2].[payment_amount], [Q2].[monthly_amount], [Q2].[total_amount], [Q2].[annual_amount], " & _
                                "[Q3].[year] AS [BOID_year], [Q3].[BO_ID] " & _
                                "INTO [" & TMP24 & "] " & _
                                "FROM ([" & TMP21 & "] AS Q1 LEFT JOIN [" & TMP22 & "] AS Q2 ON Q1.[ID] = Q2.[CCM_ID]) " & _
                                "LEFT JOIN [" & TMP23 & "] AS [Q3] ON [Q1].[ID] = [Q3].[CCM_ID] " & _
                                "ORDER BY Q1.[ID];"
                
                If f_RunQuery(stSQL5) = True Then
                        stSQL4 = ""
                        If f_Exist_LocalTable(TMP24) = True Then
                                If f_ExportExcel(TMP24, fileName, Format(Now(), "yyyymmddhhmm")) = True Then
                                        f_Download = True
                                Else
                                        ANS = MsgBox("Export failed." & vbCrLf & "Please ask administrator to check the data.", vbCritical + vbOKOnly)
                                        f_Download = False
                                End If
                        End If
                End If
        Else
                ANS = MsgBox("Something went wrong, the data might have error." & vbCrLf & _
                                        "Please ask administrator to check the data.", vbCritical + vbOKOnly)
                f_Download = False
        End If
        
        'Delete Temporary Tables ===========================================================================================
        Call f_Delete_LocalTable(TMP21)
        Call f_Delete_LocalTable(TMP22)
        Call f_Delete_LocalTable(TMP23)
        Call f_Delete_LocalTable(TMP24)
        
ERR:
        DoCmd.Hourglass False
End Function
