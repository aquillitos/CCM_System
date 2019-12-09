Attribute VB_Name = "61_Master"
Option Compare Database:    Option Explicit

Dim sDate As Double, eDate As Double
Dim RPT01 As String, RPT02 As String, RPT03 As String, RPT04 As String, RPT05 As String, RPT06 As String
        
Public Function f_Export_Report(sDateO, eDateO)
        sDate = sDateO
        eDate = eDateO
        On Error GoTo ERR_EXIT
        Dim excName As String, excNameOn As String
        Dim sheName As String
        Dim stSQL1 As String
        Dim rs1 As DAO.Recordset
        
        RPT01 = "tmp_Report01"
        RPT02 = "tmp_Report02"
        RPT03 = "tmp_Report03"
        RPT04 = "tmp_Report04"
        RPT05 = "tmp_Report05"
        RPT06 = "tmp_Report06"
        
        stSQL1 = "SELECT [value] FROM " & MST00 & " WHERE [ID] = 5 ;"
        Set rs1 = db.OpenRecordset(stSQL1, dbReadOnly)
        If Not rs1.EOF Then
                excNameOn = rs1![Value] & "_" & Format(Now(), "yyyymmddhhmm") & ".xlsx"
                excName = DeskTopPath & "\" & excNameOn
        End If
        rs1.Close: Set rs1 = Nothing:   stSQL1 = ""
        sheName = "CCM_Report"
        
        If f_Create_Pivot_Table() = True Then
                If f_ExportExcel(RPT01, excName, sheName) = True Then
                        ANS = MsgBox("The Report has been Exported on your desktop." & vbCrLf & _
                                                 excNameOn, vbInformation + vbOKOnly)
                Else
ERR_EXIT:
                        ANS = MsgBox("Error occurred on exporting excel." & vbCrLf & _
                                                "Please ask administrator for the error.", vbCritical + vbOKOnly)
                End If
        Else
                ANS = MsgBox("Error occurred on creating table." & vbCrLf & _
                                                "Please ask administrator for the error.", vbCritical + vbOKOnly)
        End If
        
        RPT01 = "": RPT02 = "": RPT03 = ""
        sDate = 0:  eDate = 0
End Function

Public Function f_Create_Pivot_Table() As Boolean
        f_Create_Pivot_Table = False
        Dim stSQL1 As String, stSQL2 As String, stSQL3 As String, stSQL4 As String
        Dim stSQL5 As String, rs5 As DAO.Recordset
        Dim m As Integer, mVar As Integer
        Dim rDate As Double, tAmount As Double
        
        On Error GoTo ERR_EXIT
        
        '期間取得 =============================================================================
        mVar = DateDiff("m", sDate, eDate)
        
        'Contract: SQL DB から Localに、CCMDATAをコピーする =========================================
        Call f_Delete_LocalTable(RPT01)
        stSQL1 = "SELECT * INTO [" & RPT01 & "] FROM " & ConSQL & ".[CCM." & CCMDATA & "];"
        'ConLo.Execute stSQL1:   stSQL1 = ""
        Call f_RunQuery(stSQL1):    stSQL1 = ""
        
        'COST: SQL DB から Localに、CCMDATAをコピーする ===========================================
        Call f_Delete_LocalTable(RPT02)
        stSQL2 = "SELECT [Q1].* INTO [" & RPT02 & "] FROM " & ConSQL & ".[CCM." & CCMCOST1 & "] AS [Q1] INNER JOIN " & _
                    "(SELECT [CCM_ID], Max([ID]) AS [ID_Max] " & _
                    "FROM " & ConSQL & ".[CCM." & CCMCOST1 & "] GROUP BY [CCM_ID] ) AS [Q2] " & _
                    "ON [Q1].[ID] = [Q2].[ID_Max] AND [Q1].[CCM_ID] = [Q2].[CCM_ID];"
        Call f_RunQuery(stSQL2):    stSQL2 = ""
        
        'BOID: SQL DB から Localに、CCMDATAをコピーする ===========================================
        Call f_Delete_LocalTable(RPT03)
        stSQL3 = "SELECT [Q1].* INTO [" & RPT03 & "] FROM " & ConSQL & ".[CCM." & CCMBOID & "] AS [Q1] INNER JOIN " & _
                        "(SELECT MAX([year]) AS [Year_Max] FROM " & ConSQL & ".[CCM." & CCMBOID & "]) AS [Q2] " & _
                        "ON [Q1].[year] = [Q2].[Year_Max] ;"
        Call f_RunQuery(stSQL3):    stSQL3 = ""
        
        'Contract, Cost, BOID をマージする ==========================================================
         If f_Exist_LocalTable(RPT01) = True And f_Exist_LocalTable(RPT02) = True Then
                stSQL4 = "SELECT [Q1].*, " & _
                                "[Q2].[terms], [Q2].[currency], [Q2].[payment_schedule], " & _
                                "[Q2].[payment_amount], [Q2].[monthly_amount], [Q2].[total_amount], [Q2].[annual_amount], " & _
                                "[Q3].[year] AS [BOID_year], [Q3].[BO_ID] " & _
                                "INTO [" & RPT04 & "] " & _
                                "FROM ([" & RPT01 & "] AS Q1 LEFT JOIN [" & RPT02 & "] AS Q2 ON Q1.[ID] = Q2.[CCM_ID]) " & _
                                "LEFT JOIN [" & RPT03 & "] AS [Q3] ON [Q1].[ID] = [Q3].[CCM_ID] " & _
                                "ORDER BY Q1.[ID];"
                Call f_RunQuery(stSQL4)
        End If
        
        'COST: SQL DB から Localに、CCMDATAをコピーする ===========================================
        Call f_Delete_LocalTable(RPT05)
        stSQL5 = "SELECT * INTO [" & RPT05 & "] FROM " & ConSQL & ".[CCM.RPT_Monthly_Cost] " & _
                        "WHERE [month] >= " & sDate & " AND [month] <= " & eDate & ";"
        ConLo.Execute stSQL5:   stSQL5 = ""
        
        
        '選択した月数分繰り返す ==================================================================
        For m = 0 To mVar
                '繰り返し毎に月を増やしていく ------------------------------------------------------------------------------------------
                rDate = DateAdd("m", m, sDate)
                
                'コピーしたCCMDATAに、料金データを月毎Joinしていく ----------------------------------------------------
                stSQL3 = "SELECT [Q1].*, IIF([Q2].[cost] IS NULL, 0, [Q2].[cost]) AS [" & Format(rDate, "YYYY-MM") & "] " & _
                            "INTO [" & RPT06 & "] FROM [" & RPT04 & "] Q1 " & _
                            "LEFT JOIN (SELECT * FROM [" & RPT05 & "] WHERE [month] = " & rDate & ") AS Q2 " & _
                            "ON [Q1].[ID] = [Q2].[CCM_ID] ;"
                
                If f_RunQuery(stSQL3) = True Then
                        '繰り返しにより同じ名前のテーブルを利用する為、コピー前のテーブルを削除して、コピー後のテーブル名をコピー前のテーブル名に書き換える。
                        'Delete Original Table
                        Call f_Delete_LocalTable(RPT04)
                        'Copy Joined Table to Original Name
                        DoCmd.Rename RPT06, acTable, RPT04
                        stSQL3 = ""
                End If
        Next m
        m = 0
        
        '期間合計値を計算する ====================================================================
        stSQL4 = "ALTER TABLE [" & RPT04 & "] ADD [Total] DOUBLE ;" '合計用Fieldを追加 --------------------------
        If f_RunQuery(stSQL4) = True Then
                stSQL4 = ""
                
                stSQL5 = "SELECT * FROM [" & RPT04 & "]"
                Set rs5 = db.OpenRecordset(stSQL5, dbOpenDynaset)
                        Do Until rs5.EOF
                                tAmount = 0
                                For m = 0 To mVar
                                        rDate = DateAdd("m", m, sDate)
                                        tAmount = tAmount + rs5(Format(rDate, "YYYY-MM"))
                                Next m
                                rs5.Edit
                                        rs5![Total] = tAmount
                                rs5.Update
                                rs5.MoveNext
                        Loop
                rs5.Close:  Set rs5 = Nothing:  stSQL5 = ""
                tAmount = 0
        Else
                ANS = MsgBox("Error on Calculating Period Total Cost.", vbInformation + vbOKOnly)
        End If
        
        mVar = 0
        Call f_Delete_LocalTable(RPT05)
        f_Create_Pivot_Table = True:    Exit Function
        
ERR_EXIT:
        f_Create_Pivot_Table = False
End Function

Public Function f_Create_SQL_String()
        Dim stSQL As String, stSQL_SET As String
        Dim stSQL_WH As String, stSQL_WH1 As String, stSQL_WH2 As String, stSQL_WH3 As String
        
        Dim stSQLSelect As String
        
        With Forms(FRM05)
        
                stSQL = "UPDATE [" & .Combo301 & "] "
                stSQLSelect = "SELECT COUNT([ID]) AS [IDS] FROM [" & .Combo301 & "] "
        
                If IsNull(.Text351) = False Then
                        If IsNumeric(.Text351) = True Then
                                stSQL_SET = "SET [" & .Combo351 & "] = " & .Text351 & " "
                        ElseIf IsDate(.Text351) = True Then
                                stSQL_SET = "SET [" & .Combo351 & "] = '" & Format(.Text351, "yyyy/mm/dd") & "' "
                        Else
                                stSQL_SET = "SET [" & .Combo351 & "] = '" & .Text351 & "' "
                        End If
                Else
                        Exit Function
                End If
        
                stSQL = stSQL & stSQL_SET
                        
                If Nz(.Combo311, "") <> "" And Nz(.Combo321, "") <> "" Then
                        stSQL_WH1 = "[" & .Combo311 & "] " & .Combo321
                        If IsNumeric(.Text311) = True Then
                                stSQL_WH1 = stSQL_WH1 & " " & .Text311 & " "
                        ElseIf IsDate(.Text311) = True Then
                                stSQL_WH1 = stSQL_WH1 & " '" & Format(.Text311, "yyyy/mm/dd") & "' "
                        ElseIf .Text311 = "null" Then
                                stSQL_WH1 = stSQL_WH1 & " NULL "
                        ElseIf .Text311 = "not null" Then
                                stSQL_WH1 = stSQL_WH1 & " NOT NULL "
                        Else
                                stSQL_WH1 = stSQL_WH1 & " '" & .Text311 & "' "
                        End If
                End If
        
                If Nz(.Combo312, "") <> "" And Nz(.Combo322, "") <> "" Then
                        stSQL_WH2 = "[" & .Combo312 & "] " & .Combo322
                        If IsNumeric(.Text312) = True Then
                                stSQL_WH2 = stSQL_WH2 & " " & .Text312 & " "
                        ElseIf IsDate(.Text312) = True Then
                                stSQL_WH2 = stSQL_WH2 & " '" & Format(.Text312, "yyyy/mm/dd") & "' "
                        ElseIf .Text312 = "null" Then
                                stSQL_WH2 = stSQL_WH2 & " NULL "
                        ElseIf .Text312 = "not null" Then
                                stSQL_WH2 = stSQL_WH2 & " NOT NULL "
                        Else
                                stSQL_WH2 = stSQL_WH2 & " '" & .Text312 & "' "
                        End If
                End If
        
                If Nz(.Combo313, "") <> "" And Nz(.Combo323, "") <> "" Then
                        stSQL_WH3 = "[" & .Combo313 & "] " & .Combo323
                        If IsNumeric(.Text313) = True Then
                                stSQL_WH3 = stSQL_WH3 & " " & .Text313 & " "
                        ElseIf IsDate(.Text313) = True Then
                                stSQL_WH3 = stSQL_WH3 & " '" & Format(.Text313, "yyyy/mm/dd") & "' "
                        ElseIf .Text313 = "null" Then
                                stSQL_WH3 = stSQL_WH3 & " NULL "
                        ElseIf .Text313 = "not null" Then
                                stSQL_WH3 = stSQL_WH3 & " NOT NULL "
                        Else
                                stSQL_WH3 = stSQL_WH3 & " '" & .Text313 & "' "
                        End If
                End If
        
                stSQL_WH = ""
                If stSQL_WH1 <> "" Then
                        stSQL_WH = "WHERE " & stSQL_WH1
                End If
                If stSQL_WH2 <> "" Then
                        If stSQL_WH = "" Then
                                stSQL_WH = " WHERE " & stSQL_WH2
                        Else
                                stSQL_WH = stSQL_WH & "AND " & stSQL_WH2
                        End If
                End If
                If stSQL_WH3 <> "" Then
                        If stSQL_WH = "" Then
                                stSQL_WH = " WHERE " & stSQL_WH3
                        Else
                                stSQL_WH = stSQL_WH & "AND " & stSQL_WH3
                        End If
                End If
        
                stSQL = stSQL & stSQL_WH
                
                .Text361 = stSQL
                .Text362 = stSQLSelect & stSQL_WH
        End With
End Function

Public Function f_Check_NumberOfRecord(stSQL As String) As String
        f_Check_NumberOfRecord = 0
        On Error GoTo ERR_EXIT
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        
        stSQL = Replace(stSQL, "FROM [", "FROM CCM.[", 1, 1)
        rs.Open stSQL, ConSys, adOpenForwardOnly, adLockReadOnly
        If Not rs.EOF Then
                f_Check_NumberOfRecord = rs![IDS]
        Else
                f_Check_NumberOfRecord = 0
        End If
        rs.Close:  Set rs = Nothing:  stSQL = ""
        Exit Function
        
ERR_EXIT:
        'SQLにエラーがある場合 ----------------------------------------------------------
        f_Check_NumberOfRecord = "*** Error ***"
        'rs.Close:  Set rs = Nothing:  stSQL = ""
End Function

Public Function f_Run_UpdateQuery(stSQL As String)
        On Error GoTo ERR_EXIT
        stSQL = Replace(stSQL, "[", "CCM.[", 1, 1)
        ConSys.Execute stSQL:  stSQL = ""
        ANS = MsgBox("Updated.", vbInformation + vbOKOnly)
        Exit Function
        
ERR_EXIT:
        ANS = MsgBox("Not Updated, Error occurred on SQL update.", vbInformation)
End Function


Public Function f_Bulk_Upload(tName As String) As Boolean
        f_Bulk_Upload = False
        
        '===== エクセルファイルインポート =================================================================
        Const IMPUP As String = "IMP_Upload"
        'UploadするExcelのパスを取得
        If f_File_PickUp("Bulk Upload", IMPUP) = False Then
                ANS = MsgBox("Error on filepath." & vbCrLf & "Check if the file is existing.", vbCritical)
                f_Bulk_Upload = False:  Exit Function
        End If
        
        'Excelをインポート
        If f_Import_Data(IMPUP, filePath) = False Then
                ANS = MsgBox("Error on importing." & vbCrLf & "Check if the file is existing.", vbCritical)
                f_Bulk_Upload = False:  Exit Function
        End If
        
        '===== データチェック ===========================================================================
        Dim rs1 As DAO.Recordset, rs2 As ADODB.Recordset
        Dim rsField1 As Field, rsField2 As ADODB.Field
        Dim c As Integer, i As Integer
        Dim fNames As String
        Dim IDCheck As Boolean
        Dim upFields() As Variant
        
        c = 0:  i = 0:  fNames = ""
        IDCheck = False
        
        'Importしたデータのフィールド名が、Upadteされるデータのフィールド名と同じであるかチェック
        Set rs1 = db.OpenRecordset(IMPUP, dbReadOnly)
        Set rs2 = New ADODB.Recordset
        rs2.Open tName, ConSys, adOpenForwardOnly, adLockReadOnly
        
        'Importしたデータのフィールド名を一つずつ取得 -------------------------------------------------------------------------------------------
        For Each rsField1 In rs1.Fields
                If rsField1.Name = "ID" Then
                        'IDをKeyにしてUpdateするので、IDフィールドが入っているか確認。
                        IDCheck = True
                Else
                        'ID以外は、Updateクエリに入れるため、Arrayに格納
                        ReDim Preserve upFields(i)
                        upFields(i) = rsField1.Name
                        i = i + 1
                End If
                
                For Each rsField2 In rs2.Fields
                        'ImportのFieldと同じFieldが、Update先のテーブルにあるかチェック。
                        If rsField2.Name = rsField1.Name Then
                                '存在すれば、次のFieldのチェックに進む
                                GoTo NEXT_Field
                        End If
                Next
                '存在しない場合書き出し用に保存
                c = c + 1:  fNames = IIf(fNames = "", rsField1.Name, fNames & ", " & rsField1.Name)
NEXT_Field:

        Next
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Debug.Print c & ": " & fNames
        
        rs1.Close:  Set rs1 = Nothing:  rs2.Close:  Set rs2 = Nothing
        Set rsField1 = Nothing: Set rsField2 = Nothing

        'ID field がなかったら中止 --------------------------------------------------------------------------------------------------------------------------------
        If IDCheck = False Then
                ANS = MsgBox("Error, aborted." & vbCrLf & "Import data must have the [ID] field.", vbCritical + vbOKOnly) '
                Erase upFields
                f_Bulk_Upload = False:  Exit Function
        End If
        '一つでもImportしたFieldが、Update先テーブルに存在しない場合中止------------------------------------------------------------------
        If c > 0 Then
                ANS = MsgBox("Error, aborted." & vbCrLf & c & " of the fields are not exist in the data." & vbCrLf & _
                                        "     Field: " & fNames, vbCritical + vbOKOnly)
                f_Bulk_Upload = False:  Exit Function
        End If
        c = 0:  i = 0
        fNames = "":    IDCheck = False:
        
        '===== バルクアップデート =======================================================================
        Dim stSQL3 As String, rs3 As DAO.Recordset
        Dim stSQL4 As String, rs4 As ADODB.Recordset
        Set rs4 = New ADODB.Recordset
        Dim fieldsString As String
        
        'UpdateすべきFieldの取得 ---------------
        For i = 0 To UBound(upFields)
                If fieldsString = "" Then
                        fieldsString = "[" & upFields(i) & "]"
                Else
                        fieldsString = fieldsString & ", [" & upFields(i) & "]"
                End If
        Next i
        
        'Update ------------------------------------------
        stSQL3 = "SELECT * FROM [" & IMPUP & "]"
        Set rs3 = db.OpenRecordset(stSQL3, dbReadOnly)
        Do Until rs3.EOF
                stSQL4 = "SELECT " & fieldsString & " FROM CCM.[" & tName & "] WHERE [ID] = " & Nz(rs3![ID], 0) & " ;"
                rs4.Open stSQL4, ConSys, adOpenDynamic, adLockOptimistic
                        If Not rs4.EOF Then
                                For i = 0 To UBound(upFields)
                                        rs4(upFields(i)) = rs3(upFields(i))
                                Next
                                rs4.Update
                        Else
                                rs4.AddNew
                                For i = 0 To UBound(upFields)
                                        rs4(upFields(i)) = rs3(upFields(i))
                                Next
                                rs4.Update
                        End If
                rs4.Close: stSQL4 = ""
                rs3.MoveNext
        Loop
        
        Erase upFields
        Set rs4 = Nothing
        rs3.Close:  Set rs3 = Nothing:  stSQL3 = ""
        
        ANS = MsgBox("Upload Completed.", vbInformation)
        f_Bulk_Upload = True
End Function

Public Function f_Upsert(tName As String) As Boolean
        f_Upsert = False
        On Error GoTo ERR_EXIT
        Dim rs1 As DAO.Recordset
        Dim stSQL2 As String, rs2 As DAO.Recordset
        Dim stSQL3 As String, rs3 As ADODB.Recordset
        Dim stSQL4 As String, rs4 As ADODB.Recordset
        Dim stSQL5 As String, rs5 As DAO.Recordset
        Dim rs1Field As DAO.Field
        
        Dim fArray() As Variant
        Dim i As Integer, j As Integer, k As Integer
        
        'Get Field Names ----------------------------------------------------------------------------------------------------------------------
        i = 0
        Set rs1 = db.OpenRecordset(tName, dbReadOnly)
        For Each rs1Field In rs1.Fields
                If Not rs1Field.Name = "ID" Then
                        ReDim Preserve fArray(i)
                        fArray(i) = rs1Field.Name
                        i = i + 1
                End If
        Next
        rs1.Close:  Set rs1 = Nothing:  Set rs1Field = Nothing
        i = 0
        
        'Update and Insert ---------------------------------------------------------------------------------------------------------------------
        Set rs3 = New ADODB.Recordset
        stSQL2 = "SELECT * FROM [" & tName & "] ORDER BY [ID];"
        Set rs2 = db.OpenRecordset(stSQL2, dbReadOnly)
                Do Until rs2.EOF
                        stSQL3 = "SELECT * FROM [CCM].[" & tName & "] WHERE [ID] = " & rs2![ID] & ";"
                        rs3.Open stSQL3, ConSys, adOpenDynamic, adLockOptimistic
                                If rs3.EOF Then
                                        rs3.AddNew
                                End If
                                        For j = 0 To UBound(fArray)
                                                rs3(fArray(j)) = rs2(fArray(j))
                                        Next
                                rs3.Update
                        rs3.Close:  stSQL3 = ""
                        rs2.MoveNext
                Loop
        Set rs3 = Nothing
        rs2.Close: Set rs2 = Nothing:  stSQL2 = ""
        j = 0
     
        'Delete -------------------------------------------------------------------------------------------------------------------------------------
        Set rs4 = New ADODB.Recordset
        stSQL4 = "SELECT [ID] FROM [" & tName & "] ORDER BY [ID] ;"
        rs4.Open stSQL4, ConSys, adOpenDynamic, adLockOptimistic
                Do Until rs4.EOF
                        stSQL5 = "SELECT [ID] FROM [" & tName & "] WHERE [ID] = " & rs4![ID] & ";"
                        Set rs5 = db.OpenRecordset(stSQL5, dbReadOnly)
                                If rs5.EOF Then
                                        rs4.Delete
                                End If
                        rs5.Close:  stSQL5 = ""
                        rs4.MoveNext
                Loop
        Set rs5 = Nothing
        rs4.Close:  Set rs4 = Nothing:  stSQL4 = ""
        
        f_Upsert = True
        Exit Function
        
ERR_EXIT:
        f_Upsert = False
End Function
