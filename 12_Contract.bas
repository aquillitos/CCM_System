Attribute VB_Name = "12_Contract"
Option Compare Database:    Option Explicit
    
 Public Function f_Choose_Item(cID As Long)
        Dim rs1 As ADODB.Recordset, stSQL1 As String
        Dim i As Integer, j As Integer, fName As String, cName As String
        
        Set rs1 = New ADODB.Recordset
        stSQL1 = "SELECT * FROM CCM.[" & CCMDATA & "] WHERE ID = " & cID & ";"
        rs1.Open stSQL1, ConSys, adOpenForwardOnly, adLockReadOnly
        
        With Forms(FRM03)
                If Not rs1.EOF Then
                        For i = 1 To 32
                            cName = "Text" & 100 + i
                            fName = .Controls(cName).Properties("ControlTipText")
                            .Controls(cName) = rs1(fName)
                        Next i
                        For j = 1 To 5
                            cName = "Text" & 150 + j
                            fName = .Controls(cName).Properties("ControlTipText")
                            .Controls(cName) = rs1(fName)
                        Next j
                        
                        If IsDate(.Text104) = True And IsDate(.Text106) = True Then
                                Dim sMonth As Date
                                sMonth = DateSerial(Year(DateAdd("m", -1, .Text104)), Month(DateAdd("m", -1, .Text104)), 1)
                                .Controls("Text133") = DateDiff("m", sMonth, .Text106) & " Months"
                        End If
                        
                        contractID = rs1![ID]
                        contractNumber = rs1![number]
                        contractStatus = rs1![status]
                        cName = "": fName = ""
                End If
                Call f_Check_Item_Cancel(contractStatus)
        End With
        rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
        i = 0:  j = 0
        Call f_Choose_BOID(contractNumber)
End Function

Public Function f_Choose_BOID(cID As String)
        Dim stSQL1 As String, rs1 As ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        stSQL1 = "SELECT TOP 1 [year], [BO_ID] FROM CCM.[" & CCMBOID & "] WHERE [CCM_number] = '" & cID & "' ORDER BY [year], [BO_ID];"
        rs1.Open stSQL1, ConSys, adOpenForwardOnly, adLockReadOnly
        With Forms(FRM03)
                If Not rs1.EOF Then
                        .Text141 = rs1![Year]
                        .Text142 = rs1![BO_ID]
                Else
                        .Text141 = ""
                        .Text142 = ""
                End If
        End With
        rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
End Function

Public Function f_Check_Item_Cancel(status As String)
        Select Case status
                Case Is = "Active"
                        Call f_Control_Item_Locked(0)
                Case Is = "Cancelled"
                        Call f_Control_Item_Locked(1)
                Case Is = "will be Cancelled"
                        Call f_Control_Item_Locked(1)
                Case Else
                        Call f_Control_Item_Locked(0)
        End Select
End Function

Public Function f_Control_Item_Locked(s As Integer) As Integer
        On Error Resume Next
        Dim nTag As Integer:    Dim meControl As Control
        
        With Forms(FRM03)
        For Each meControl In .Controls
                If meControl.Name Like "Text1*" Then
                        nTag = Nz(meControl.Properties("tag"), 0)
                        Select Case s
                                Case 0 'ActiveÇ≈ëSÇƒLocked
                                        If nTag >= 0 Then
                                            meControl.Locked = True
                                            meControl.BackColor = RGB(236, 236, 236)
                                        End If
                                Case 1 'CancelledÇ≈ëSÇƒLocked
                                        If nTag >= 0 Then
                                                meControl.Locked = True
                                                meControl.BackColor = RGB(200, 200, 200)
                                                .ComContractEdit.Enabled = False
                                                .ComPriceEdit.Enabled = False
                                                .ComPriceNew.Enabled = False
                                                .ComAttachDelete.Enabled = False
                                                .ComAttachAdd.Enabled = False
                                        End If
                                Case 2 'ActiveÇ≈Lockâèú
                                        If nTag < 5 Then
                                                meControl.Locked = False
                                                meControl.BackColor = RGB(255, 255, 255)
                                        End If
                                Case 3
                                        If nTag < 0 Then
                                                meControl.Locked = True
                                                meControl.BackColor = RGB(236, 236, 236)
                                        End If
                                Case 5
                        End Select
                End If
        Next
        Set meControl = Nothing
        End With
        On Error GoTo 0
End Function

Public Function f_Update_Check(cStatus As String) As Boolean
        f_Update_Check = False
        Dim cArray As Variant
        Dim cName As String
        Dim t As Integer, i As Integer, c As Integer
        c = 0
        
        Select Case cStatus
                Case "Active"
                        cArray = Array(2, 3, 4, 8, 10, 11, 12, 13, 14, 15, 16, 19, 23, 27, 29, 31, 32, 41, 42)
                Case "Draft"
                        cArray = Array(2, 3, 8, 12, 14, 23)
                Case "Cancelled"
                        cArray = Array(2, 3, 7)
                Case Else
                        cArray = Array(2, 3)
        End Select
                
        With Forms(FRM03)
                On Error Resume Next
                For t = 1 To 55
                        cName = "Text" & 100 + t
                        If .Controls(cName).Locked = False Then
                                .Controls(cName).BackColor = RGB(255, 255, 255)
                        End If
                Next t
                For i = 0 To UBound(cArray)
                        cName = "Text" & 100 + cArray(i)
                        If Nz(Trim(.Controls(cName)), "") = "" Then
                                .Controls(cName).BackColor = RGB(239, 211, 210)
                                c = c + 1
                        Else
                                .Controls(cName).BackColor = RGB(255, 255, 255)
                        End If
                Next i
                On Error GoTo 0
        End With
        If c > 0 Then
                f_Update_Check = False
        Else
                f_Update_Check = True
        End If
        
        Erase cArray
End Function

Public Function f_Update_CCM() As Boolean
        f_Update_CCM = False
        Dim stSQL1 As String, rs1 As ADODB.Recordset
        Dim i As Integer, j As Integer, cName As String, fName As String
        
        Set rs1 = New ADODB.Recordset
        stSQL1 = "SELECT * FROM CCM.[" & CCMDATA & "] WHERE [ID]  = " & contractID & ";"
        rs1.Open stSQL1, ConSys, adOpenDynamic, adLockOptimistic
        If Not rs1.EOF Then
                With Forms(FRM03)
                        If Not rs1.EOF Then
                                For i = 2 To 32
                                        cName = "Text" & 100 + i
                                        fName = .Controls(cName).Properties("ControlTipText")
                                        If .Controls(cName).Properties("Tag") <= 30 Then
                                                Call f_Update_History(rs1(fName), .Controls(cName), CCMDATA, fName)
                                                rs1(fName) = .Controls(cName)
                                        End If
                                Next i
                                For j = 1 To 5
                                        cName = "Text" & 150 + j
                                        fName = .Controls(cName).Properties("ControlTipText")
                                        If .Controls(cName).Properties("Tag") <= 30 Then
                                                Call f_Update_History(rs1(fName), .Controls(cName), CCMDATA, fName)
                                                rs1(fName) = .Controls(cName)
                                        End If
                                Next j
                                rs1.Update
                                cName = "": fName = ""
                        End If
                End With
                f_Update_CCM = True
        Else
                f_Update_CCM = False
        End If
        rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
End Function

Public Function f_Update_BOID(bYear As Integer) As Boolean
        f_Update_BOID = False
        On Error GoTo ERR
        Dim stSQL1 As String, rs1 As ADODB.Recordset
        Dim i As Integer, cName As String, fName As String
        
        Set rs1 = New ADODB.Recordset
        
        stSQL1 = "SELECT * FROM CCM.[" & CCMBOID & "] WHERE [CCM_ID]  = " & contractID & " AND [year] = " & bYear & ";"
        rs1.Open stSQL1, ConSys, adOpenDynamic, adLockOptimistic
        
        If rs1.EOF Then
                rs1.AddNew
                rs1![CCM_ID] = contractID
                rs1![CCM_number] = contractNumber
        End If
        
        With Forms(FRM03)
                For i = 41 To 42
                        cName = "Text" & 100 + i
                        fName = .Controls(cName).Properties("ControlTipText")
                        If .Controls(cName).Properties("Tag") <= 30 Then
                                If rs1.EOF Then
                                        Call f_Update_History("*new entry", .Controls(cName), CCMDATA, fName)
                                Else
                                        Call f_Update_History(rs1(fName), .Controls(cName), CCMDATA, fName)
                                End If
                        End If
                        rs1(fName) = .Controls(cName)
                        cName = "": fName = ""
                Next i
        End With
        
        rs1.Update
        f_Update_BOID = True
        rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
        Exit Function
        
ERR:
        f_Update_BOID = False
End Function

Public Function f_New_CCM() As Boolean
        f_New_CCM = False
        Dim stSQL1 As String, rs1 As ADODB.Recordset
        Dim i As Integer, j As Integer, cName As String, fName As String
        Set rs1 = New ADODB.Recordset
        
        stSQL1 = "SELECT * FROM CCM.[" & CCMDATA & "];"
        rs1.Open stSQL1, ConSys, adOpenDynamic, adLockOptimistic
        If Not rs1.EOF Then
                With Forms(FRM03)
                        rs1.AddNew
                        For i = 2 To 32
                                cName = "Text" & 100 + i
                                fName = .Controls(cName).Properties("ControlTipText")
                                If Nz(.Controls(cName), "") <> "" Then
                                        Call f_Update_History("*new entry", .Controls(cName), CCMDATA, fName)
                                        rs1(fName) = .Controls(cName)
                                End If
                        Next i
                        For j = 1 To 5
                                cName = "Text" & 150 + j
                                fName = .Controls(cName).Properties("ControlTipText")
                                If Nz(.Controls(cName), "") <> "" Then
                                        Call f_Update_History("*new entry", .Controls(cName), CCMDATA, fName)
                                        rs1(fName) = .Controls(cName)
                                End If
                        Next j
                        rs1.Update
                        cName = "": fName = ""
                End With
                f_New_CCM = True
        Else
                f_New_CCM = False
        End If
        rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
End Function

Public Function f_New_BOID(bYear As Integer) As Boolean
        f_New_BOID = False
        Dim stSQL1 As String, rs1 As ADODB.Recordset
        Dim i As Integer, cName As String, fName As String
        Set rs1 = New ADODB.Recordset
        
        stSQL1 = "SELECT * FROM CCM.[" & CCMBOID & "];"
        rs1.Open stSQL1, ConSys, adOpenDynamic, adLockOptimistic
        If Not rs1.EOF Then
                With Forms(FRM03)
                        rs1.AddNew
                        rs1![CCM_ID] = contractID
                        rs1![CCM_number] = contractNumber
                        For i = 41 To 42
                                cName = "Text" & 100 + i
                                fName = .Controls(cName).Properties("ControlTipText")
                                If Nz(.Controls(cName), "") <> "" Then
                                        Call f_Update_History("*new entry", .Controls(cName), CCMDATA, fName)
                                        rs1(fName) = .Controls(cName)
                                End If
                        Next i
                        rs1.Update
                        cName = "": fName = ""
                End With
                f_New_BOID = True
        Else
                f_New_BOID = False
        End If
        rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
End Function

Public Function f_Get_New_ID(cNumber As String) As Boolean
        f_Get_New_ID = False
        Dim stSQL1 As String, rs1 As ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        stSQL1 = "SELECT [ID], [number], [status] FROM CCM.[" & CCMDATA & "] WHERE [number] = '" & cNumber & "';"
        rs1.Open stSQL1, ConSys, adOpenDynamic, adLockReadOnly
                If Not rs1.EOF Then
                        contractID = rs1![ID]
                        contractNumber = rs1![number]
                        contractStatus = rs1![status]
                        f_Get_New_ID = True
                Else
                        contractID = 0
                        f_Get_New_ID = False
                        ANS = MsgBox("Error occurred." & vbCrLf & "Please ask Administrator.", vbCritical + vbOKOnly)
                End If
        rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
End Function

Public Function f_Button_Unlock_Item(cStatus) As Integer
        Dim stSQL1 As String, rs1 As ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        f_Button_Unlock_Item = 1
        stSQL1 = "SELECT * FROM CCM.[" & MST03 & "] WHERE status = '" & cStatus & "';"
        rs1.Open stSQL1, ConSys, adOpenDynamic, adLockReadOnly
                If Not rs1.EOF Then
                        f_Button_Unlock_Item = rs1![initial_lock]
                End If
        rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
End Function

Public Function f_Check_Change_Status(cNumber As String, cStatus As String) As Integer
        f_Check_Change_Status = 0
        Dim stSQL1 As String, stSQL2 As String, stSQL3 As String
        Dim priorityFrom As Integer, priorityTo As Integer, noChange As Integer
        Dim rs1 As ADODB.Recordset, rs2 As ADODB.Recordset, rs3 As ADODB.Recordset
        
        priorityFrom = 0:   priorityTo = 100:   noChange = 1
        
        Set rs1 = New ADODB.Recordset
        Set rs2 = New ADODB.Recordset
        Set rs3 = New ADODB.Recordset
        
        stSQL1 = "SELECT [status] FROM CCM.[" & CCMDATA & "] WHERE [number] = '" & cNumber & "'"
        rs1.Open stSQL1, ConSys, adOpenDynamic, adLockReadOnly
        If Not rs1.EOF Then
                stSQL2 = "SELECT [priority] FROM CCM.[" & MST03 & "] WHERE [status] = '" & rs1![status] & "';"
                rs2.Open stSQL2, ConSys, adOpenDynamic, adLockReadOnly
                If Not rs2.EOF Then
                        priorityFrom = rs2![Priority]
                End If
                rs2.Close:  Set rs2 = Nothing:  stSQL2 = ""
        End If
        rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
        
        stSQL3 = "SELECT [priority], [no_change_update] FROM CCM.[" & MST03 & "] WHERE [status] = '" & cStatus & "';"
        rs3.Open stSQL3, ConSys, adOpenDynamic, adLockReadOnly
                If Not rs3.EOF Then
                        priorityTo = rs3![Priority]
                        noChange = rs3![no_change_update]
                End If
        rs3.Close:  Set rs3 = Nothing:  stSQL3 = ""
        
        Select Case priorityTo - priorityFrom
            Case Is < 0
                    f_Check_Change_Status = 0
            Case Is = 0
                    Select Case noChange
                            Case 0
                                    f_Check_Change_Status = 1
                            Case 1
                                    f_Check_Change_Status = 2
                    End Select
            Case Is > 0
                    Select Case noChange
                            Case 0
                                    f_Check_Change_Status = 3
                            Case Else
                                    f_Check_Change_Status = 4
                    End Select
        End Select
        priorityTo = 0: priorityFrom = 0:   noChange = 0
End Function
