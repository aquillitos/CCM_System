Attribute VB_Name = "11_Common_Function"
Option Compare Database:    Option Explicit
Dim s_Search_Query As String

Public Function f_List_Contract(fCount As Integer) As Boolean
        f_List_Contract = False
        On Error GoTo ERR
        
        Dim stSQL1 As String, stSQL2 As String
        Dim rs2 As ADODB.Recordset
        
        s_Search_Query = f_Search_Query(fCount)
        
        'åüçıåãâ Çï\é¶
        stSQL1 = "SELECT * FROM " & ConSQL & ".[CCM." & CCMDATA & "] " & s_Search_Query
        Forms(FRM01).SUB001.Form.RecordSource = stSQL1: stSQL1 = ""
        
        stSQL2 = "SELECT Count([ID]) AS IDCount FROM CCM.[" & CCMDATA & "] " & s_Search_Query
        Set rs2 = New ADODB.Recordset
        rs2.Open stSQL2, ConSys, adOpenForwardOnly, adLockReadOnly
        If rs2.EOF Then
                Forms(FRM01).Text101 = 0
        Else
                Forms(FRM01).Text101 = rs2![IDCount]
        End If
        rs2.Close:  Set rs2 = Nothing:  stSQL2 = ""
        f_List_Contract = True
        Exit Function

ERR:
        f_List_Contract = False
End Function

Public Function f_Search_Query(cCount As Integer) As String
        Dim stSQLWh As String
        Dim controlName As String, fieldName As String, searchType As Integer
        Dim i As Integer, sDate As Date, eDate As Date
        Dim sDateS As Double, eDateS As Double
        
        stSQLWh = ""
        With Forms(FRM01)
                For i = 1 To cCount 'Max Number of control
                    controlName = "Search" & Right("00" & i, 3)
                    fieldName = .Controls(controlName).Properties("DatasheetCaption").Value
                    searchType = .Controls(controlName).Properties("Tag").Value
                    If Nz(.Controls(controlName), "") <> "" Then
                            Select Case searchType
                                    Case 1 'Exact Text Search
                                            If stSQLWh = "" Then
                                                    stSQLWh = "WHERE [" & fieldName & "] = '" & .Controls(controlName) & "'"
                                            Else
                                                    stSQLWh = stSQLWh & " AND [" & fieldName & "] = '" & .Controls(controlName) & "'"
                                            End If
                                    Case 2  'Wildcard Search
                                            If stSQLWh = "" Then
                                                    stSQLWh = "WHERE [" & fieldName & "] LIKE '%" & .Controls(controlName) & "%'"
                                            Else
                                                    stSQLWh = stSQLWh & " AND [" & fieldName & "] LIKE '%" & .Controls(controlName) & "%'"
                                            End If
                                    Case 3 'Exact Date Search
                                            sDate = .Controls(controlName)
                                            sDateS = DateSerial(Year(sDate), Month(sDate), Day(sDate))
                                            If stSQLWh = "" Then
                                                    stSQLWh = "WHERE [" & fieldName & "] = " & sDate & ""
                                            Else
                                                    stSQLWh = "WHERE [" & fieldName & "] = " & sDate & ""
                                            End If
                                    Case 4 'Date Within the Month Search
                                            sDate = .Controls(controlName)
                                            eDate = DateAdd("m", 1, sDate)
                                            sDateS = DateSerial(Year(sDate), Month(sDate), Day(sDate))
                                            eDateS = DateSerial(Year(eDate), Month(eDate), Day(eDate))
                                            Debug.Print sDate & " / " & eDate
                                            If stSQLWh = "" Then
                                                    stSQLWh = "WHERE [" & fieldName & "] >= " & sDateS & " AND [" & fieldName & "] < " & eDateS & ""
                                            Else
                                                    stSQLWh = stSQLWh & " AND [" & fieldName & "] >= " & sDateS & " AND [" & fieldName & "] < " & eDateS & ""
                                            End If
                            End Select
                    End If
                Next i
                f_Search_Query = stSQLWh
        End With
End Function

Public Function f_Get_MaxNumber(numberModel As String) As String
        f_Get_MaxNumber = ""
        Dim numberPrefix As String, numLen As Integer
        Dim stSQL1 As String
        Dim rs1 As ADODB.Recordset
        
        Select Case numberModel
                Case "Circuit"
                        numberPrefix = cNumPref1
                Case "Lease"
                        numberPrefix = cNumPref2
                Case "Maintenance"
                        numberPrefix = cNumPref3
                Case "Contract"
                        numberPrefix = cNumPref4
                Case "SRV"
                        numberPrefix = cNumPref5
                Case Else
                        numberPrefix = cNumPref6
        End Select
        
        Set rs1 = New ADODB.Recordset
        stSQL1 = "SELECT [number] FROM CCM.[" & CCMDATA & "] WHERE [number] = " & _
                        "(SELECT MAX([number]) FROM CCM.[" & CCMDATA & "] WHERE [number] LIKE '" & numberPrefix & "%' );"
        rs1.Open stSQL1, ConADO, adOpenForwardOnly, adLockReadOnly
        If Not rs1.EOF Then
                numLen = Len(rs1![number]) - Len(numberPrefix)
                f_Get_MaxNumber = numberPrefix & Right("0000000" & Val(Right(rs1![number], numLen)) + 1, numLen)
        Else
                f_Get_MaxNumber = "Error"
        End If
        rs1.Close:  Set rs1 = Nothing:  stSQL1 = "":    numLen = 0
        contractNumber = f_Get_MaxNumber
End Function
