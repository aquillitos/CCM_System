Attribute VB_Name = "02_Table_Control"
Option Compare Database:    Option Explicit

Public Function f_Exist_LocalTable(tableName) As Boolean
        f_Exist_LocalTable = False
        On Error GoTo ERR_EXIT
                f_Exist_LocalTable = CurrentDb.TableDefs(tableName).Name = tableName
                f_Exist_LocalTable = True
        Exit Function
ERR_EXIT:
        f_Exist_LocalTable = False
End Function

Public Function f_Delete_LocalTable(tableName) As Boolean
        f_Delete_LocalTable = False
        On Error GoTo ERR_EXIT
                DoCmd.Close acTable, tableName: DoCmd.DeleteObject acTable, tableName:
                f_Delete_LocalTable = True
        Exit Function
ERR_EXIT:
        f_Delete_LocalTable = False
End Function

Public Function f_Exist_LocalQuery(queryName) As Boolean
        f_Exist_LocalQuery = False
        On Error GoTo ERR_EXIT
                f_Exist_LocalQuery = CurrentDb.QueryDefs(queryName).Name = queryName
                f_Exist_LocalQuery = True
        Exit Function
ERR_EXIT:
        f_Exist_LocalQuery = False
End Function

Public Function f_Delete_LocalQuery(queryName) As Boolean
        f_Delete_LocalQuery = False
        On Error GoTo ERR_EXIT
                DoCmd.Close acQuery, queryName
                DoCmd.DeleteObject acQuery, queryName
                f_Delete_LocalQuery = True
        Exit Function
ERR_EXIT:
        f_Delete_LocalQuery = False
End Function

Public Function f_Copy_TableSL(FromTable As String, ToTable As String) As Boolean  'System Database to LocaL
        f_Copy_TableSL = False
        On Error GoTo ERR_EXIT
                Dim stSQL1 As String
                Call f_Delete_LocalTable(ToTable)
                stSQL1 = "SELECT * INTO [" & ToTable & "] FROM " & ConSQL & ".[CCM." & FromTable & "]"
                ConLo.Execute stSQL1:   stSQL1 = ""
                f_Copy_TableSL = True
        Exit Function
ERR_EXIT:
        f_Copy_TableSL = False
        On Error GoTo 0
End Function

Public Function f_RunQuery(stSQL As String) As Boolean
        f_RunQuery = False
        On Error GoTo ERR_EXIT
                DoCmd.SetWarnings False
                        Debug.Print stSQL
                        DoCmd.RunSQL stSQL: stSQL = ""
                DoCmd.SetWarnings True
                f_RunQuery = True
        Exit Function
ERR_EXIT:
        f_RunQuery = False
        DoCmd.SetWarnings True
End Function

Public Function f_PT_Query(stSQL As String, queryNm As String, ReturnRec As Boolean) As Boolean
        '///// queryNm: Query Table Name created on the Access DB
        '///// ReturnRec: Choose if you need to create Local Query or Not.
        f_PT_Query = False
        Dim dbc As DAO.Database
        Dim queryDf As DAO.QueryDef
        On Error GoTo ERR_EXIT
                Set dbc = CurrentDb
                Call f_Delete_LocalQuery(queryNm)
                Set queryDf = dbc.CreateQueryDef(queryNm)
                    With queryDf
                        .Connect = Replace(Replace(ConSQL, "[", "", 1), "]", "", 1)
                        .ReturnsRecords = ReturnRec
                        .SQL = stSQL
                        .Close
                    End With
                Set queryDf = Nothing
                Set dbc = Nothing
                f_PT_Query = True
        Exit Function
ERR_EXIT:
        Set queryDf = Nothing
        Set dbc = Nothing
        f_PT_Query = Fals
End Function

Public Function f_Drop_System_Table(tableName) As Boolean
        Dim rsD As ADODB.Recordset
        Set rsD = New ADODB.Recordset
        On Error Resume Next:
        rsD.Open "DROP TABLE CCM.[" & tableName & "];", C_ConADOSys:
        rsD.Close: Set rsD = Nothing
        On Error GoTo 0
End Function

Public Function f_ExportExcel(tName As String, fName As String, sName As String) As Boolean
        f_ExportExcel = False
        Call f_DeleteFiles(fName)
        On Error GoTo ERR_EXIT:
                DoCmd.SetWarnings False
                DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, tName, fName, True, sName
                DoCmd.SetWarnings True
                f_ExportExcel = True
        Exit Function
ERR_EXIT:
        f_ExportExcel = False
End Function

Public Function f_Delete_Temporary()
        On Error Resume Next
        Dim tmpTb As TableDef
        For Each tmpTb In db.TableDefs
                Debug.Print tmpTb.Name
                If tmpTb.Name Like "tmp_*" Then
                        Call f_Delete_LocalTable(tmpTb.Name)
                End If
        Next
        Dim tmpQu As QueryDef
        For Each tmpQu In db.QueryDefs
                Debug.Print tmpQu.Name
                If tmpQu.Name Like "tmp_*" Then
                        Call f_Delete_LocalTable(tmpQu.Name)
                End If
        Next
        On Error GoTo 0
End Function

Public Function f_Delete_MasterQuery()
        Dim mstQry As QueryDef
        On Error Resume Next
        For Each mstQry In db.QueryDefs
                If mstQry.Name Like "CCM_MST_*" Then
                        Call f_Delete_LocalQuery(mstQry.Name)
                End If
        Next
End Function
