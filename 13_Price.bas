Attribute VB_Name = "13_Price"
Option Compare Database:    Option Explicit

Public Function f_List_Price(cID As Long)
        Dim stSQL1 As String, stSQL2 As String
        Forms(FRM03).Text201.RowSource = ""
        stSQL2 = "SELECT [ID], [CCM_ID], FORMAT([term_start],'yyyy-mm-dd') AS [Start], FORMAT([term_end],'yyyy-mm-dd') AS [End], [terms] AS [Term], [currency] as [CUR], " & _
                                "[payment_schedule] AS [Schedule], FORMAT([payment_amount], '#,##0') AS [Payment Amount], [memo], [locked] " & _
                                "FROM " & ConSQL & ".[CCM." & CCMCOST1 & "] WHERE [CCM_ID] = " & cID & " ORDER BY [term_start] DESC, [term_end] DESC, [ID];"
        Forms(FRM03).Text201.RowSource = stSQL2
        Call f_Check_Cancel_Price(contractStatus)
End Function

Public Function f_Choose_Price(pID As Long)
        Dim stSQL1 As String, rs1 As ADODB.Recordset
        Dim i As Integer, cName As String, fName As String
        Set rs1 = New ADODB.Recordset
        
        stSQL1 = "SELECT * FROM CCM.[" & CCMCOST1 & "] WHERE [ID] = " & pID & ";"
        rs1.Open stSQL1, ConSys, adOpenDynamic, adLockReadOnly
        With Forms(FRM03)
                If Not rs1.EOF Then
                        For i = 2 To 15
                                cName = "Text" & 200 + i
                                fName = .Controls(cName).Properties("ControlTipText")
                                .Controls(cName) = rs1(fName)
                        Next i
                End If
        End With
        rs1.Close: Set rs1 = Nothing:   stSQL1 = ""
        cName = "": fName = ""
        Call pf_Format_Price
End Function

Private Function pf_Format_Price()
        With Forms(FRM03)
                .Text208 = Format(.Text208, "##,##0")
                .Text209 = Format(.Text209, "##,##0")
                .Text210 = Format(.Text210, "##,##0")
                .Text211 = Format(.Text211, "##,##0")
        End With
End Function

Public Function f_Check_Cancel_Price(status As String) As Boolean
        Select Case status
                Case Is = "Active"
                        Call f_Control_Price_Locked(0)
                        f_Check_Cancel_Price = False
                Case Is = "Cancelled"
                        Call f_Control_Price_Locked(1)
                        f_Check_Cancel_Price = True
                Case Is = "Draft"
                        Call f_Control_Price_Locked(0)
                        f_Check_Cancel_Price = False
        End Select
End Function

Public Function f_Button_Unlock_Price(cStatus) As Integer
        f_Button_Unlock_Price = 1
        Dim stSQL1 As String, rs1 As ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        stSQL1 = "SELECT * FROM CCM.[" & MST03 & "] WHERE status = '" & cStatus & "';"
        rs1.Open stSQL1, ConSys, adOpenDynamic, adLockReadOnly
                If Not rs1.EOF Then
                        f_Button_Unlock_Price = rs1![initial_lock]
                End If
        rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
End Function

Public Function f_Control_Price_Locked(s As Integer)
        On Error Resume Next
        Dim nTag As Integer:    Dim meControl As Control
        
        With Forms(FRM03)
        For Each meControl In .Controls
                If meControl.Name Like "Text2*" Then
                        nTag = Nz(meControl.Properties("tag"), 0)
                        Select Case s
                                Case 0 'Activeで全てLocked
                                        If nTag >= 0 Then
                                            meControl.Locked = True
                                            meControl.BackColor = RGB(236, 236, 236)
                                        End If
                                Case 1 'Cancelledで全てLocked
                                        If nTag >= 0 Then
                                                meControl.Locked = True
                                                meControl.BackColor = RGB(200, 200, 200)
                                        End If
                                Case 2 'ActiveでLock解除
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
        End With
        On Error GoTo 0
End Function

Public Function f_Update_Check_Price(cStatus As String) As Boolean
        f_Update_Check_Price = 0
        Dim cArray() As Variant
        Dim cName As String
        Dim i As Integer, c As Integer
        
        If cStatus = "Active" Then
                cArray = Array(3, 4, 6, 7, 8)
        ElseIf cStatus = "Draft" Then
                cArray = Array(4)
        End If
        With Forms(FRM03)
                For i = 0 To UBound(cArray)
                        cName = "Text" & 200 + cArray(i)
                        If Nz(Trim(.Controls(cName)), "") = "" Then
                                .Controls(cName).BackColor = RGB(239, 211, 210)
                                c = c + 1
                        Else
                                If .Controls(cName).Locked = False Then
                                    .Controls(cName).BackColor = RGB(255, 255, 255)
                                End If
                        End If
                Next i
        End With
        If c > 0 Then
                f_Update_Check_Price = False
        Else
                f_Update_Check_Price = True
        End If
        Erase cArray
End Function

Public Function f_Update_Price(pID As Long) As Boolean
        f_Update_Price = False
        Dim stSQL1 As String, stSQL2 As String
        Dim rs1 As ADODB.Recordset, rs2 As ADODB.Recordset
        Dim i As Integer, cName As String
        Dim orgValue As Variant, newValue As Variant, fieldName As String
        
        On Error GoTo ERR_EXIT
        Set rs1 = New ADODB.Recordset
        Set rs2 = New ADODB.Recordset
        
        'コストリストを空にする ========================================================
        Forms(FRM03).Text201.RowSource = ""
        Forms(FRM03).Text201.Requery
        
        'Price Dataを更新する ==========================================================
        With Forms(FRM03)
                stSQL1 = "SELECT * FROM CCM.[" & CCMCOST1 & "] WHERE [ID]  = " & pID & ";"
                rs1.Open stSQL1, ConSys, adOpenDynamic, adLockOptimistic
                If Not rs1.EOF Then
                        For i = 3 To 15
                                cName = "Text" & 200 + i
                                fieldName = .Controls(cName).Properties("ControlTipText")
                                orgValue = rs1(fieldName)
                                newValue = .Controls(cName)
                                Call f_Update_History(orgValue, newValue, CCMCOST1, fieldName)
                                rs1(fieldName) = newValue
                        Next i
                        rs1.Update
                        f_Update_Price = True
                End If
                rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
        
                'CCM Data [Start] & [End] Update ==================================================
                stSQL2 = "SELECT [start], [end] FROM CCM.[" & CCMDATA & "] WHERE ID = " & contractID & ";"
                rs2.Open stSQL2, ConSys, adOpenDynamic, adLockOptimistic
                If Not rs2.EOF Then
                        Call f_Update_History(rs2![start], .Text203, CCMDATA, "start")
                        Call f_Update_History(rs2![end], .Text204, CCMDATA, "end")
                        rs2![start] = .Text203
                        rs2![end] = .Text204
                        rs2.Update
                End If
                rs2.Close:  Set rs2 = Nothing:  stSQL2 = ""
                
                'Update Monthly Price Data ======================================================
                Call pf_Update_Monthly_Cost
        End With
        
        Call f_Choose_Item(contractID)
        Call f_List_Price(contractID)
        f_Update_Price = True
        Exit Function
        
ERR_EXIT: 'エラー処理==================================================================
        If ERR.number = 3197 Then
                ANS = MsgBox("There is another person who is updating the same record." & vbCrLf & "Please wait a moment." _
                                , vbExclamation + vbOKOnly)
        Else
                ANS = MsgBox("Something went wrong, the data might have error." & vbCrLf & _
                                "Please ask administrator to check the data.", vbCritical + vbOKOnly)
        End If
        f_Update_Price = False
End Function

Public Function f_New_Price() As Boolean
        f_New_Price = False
        Dim stSQL1 As String, stSQL2 As String, stSQL3 As String, stSQL4 As String, stSQL5 As String
        Dim rs2 As ADODB.Recordset, rs3 As ADODB.Recordset, rs4 As ADODB.Recordset
        Dim i As Integer, cName As String
        Dim orgValue As Variant, newValue As Variant, fieldName As String
        
        On Error GoTo ERR_EXIT
        Set rs2 = New ADODB.Recordset
        Set rs3 = New ADODB.Recordset
        
        '既存のPrice Dataを全てLockにする ====================================================
        stSQL1 = "UPDATE CCM.[" & CCMCOST1 & "] SET [Locked] = 1 WHERE [CCM_ID] = " & contractID & ";"
        ConSys.Execute stSQL1:  stSQL1 = ""
        
        '新しいPrice Dataを追加する ==========================================================
        With Forms(FRM03)
                'Price Data Update =============================================================
                stSQL2 = "SELECT * FROM CCM.[" & CCMCOST1 & "];"
                rs2.Open stSQL2, ConSys, adOpenDynamic, adLockOptimistic
                        rs2.AddNew
                        For i = 3 To 15
                                cName = "Text" & 200 + i
                                fieldName = .Controls(cName).Properties("ControlTipText")
                                If Nz(.Controls(cName), "") <> "" Then
                                        orgValue = "*new entry"
                                        newValue = .Controls(cName)
                                        Call f_Update_History(orgValue, newValue, CCMCOST1, fieldName)
                                        rs2(fieldName) = newValue
                                End If
                        Next i
                        rs2.Update
                rs2.Close:  Set rs2 = Nothing:  stSQL2 = ""
                
                'CCM Data [Start] & [End] Update ==================================================
                stSQL3 = "SELECT * FROM CCM.[" & CCMDATA & "] WHERE ID = " & contractID & ";"
                rs3.Open stSQL3, ConSys, adOpenDynamic, adLockOptimistic
                        If Not rs3.EOF Then
                                Call f_Update_History(rs3![start], .Text203, CCMDATA, "start")
                                Call f_Update_History(rs3![end], .Text204, CCMDATA, "end")
                                rs3![start] = .Text203
                                rs3![end] = .Text204
                                rs3.Update
                        End If
                rs3.Close:  Set rs3 = Nothing:  stSQL3 = ""
                
                'Update Monthly Price Data ======================================================
                Call pf_Update_Monthly_Cost
        End With
        
        Call f_List_Price(contractID)
        f_New_Price = True
        Exit Function
          
ERR_EXIT: 'エラー処理==================================================================
        'Lock を 外す
        stSQL4 = "SELECT MAX([ID]) AS [MxID] FROM CCM.[" & CCMCOST1 & "] WHERE [CCM_ID] = " & contractID & " ;"
        Set rs4 = New ADODB.Recordset
        rs4.Open stSQL4, ConSys, adOpenDynamic, adLockOptimistic
                If Not rs4.EOF Then
                        If IsNull(rs4![MxID]) = False Then
                                stSQL5 = "UPDATE CCM.[" & CCMCOST1 & "] SET  [Locked] = 0 WHERE [ID] = " & rs4![MxID] & " "
                                ConSys.Execute stSQL5: stSQL5 = ""
                        End If
                End If
        rs4.Close:  Set rs4 = Nothing:  stSQL4 = ""
        f_New_Price = False
End Function

Public Function f_Delete_Price_Data(pID As Long)
        Dim stSQL1 As String, stSQL2 As String, stSQL3 As String
        Dim rs1 As ADODB.Recordset, rs2 As ADODB.Recordset
        
        On Error GoTo ERR_EXIT
        
        '消去処理 ===============================================================================
        stSQL1 = "DELETE FROM CCM.[" & CCMCOST1 & "] WHERE [ID] = " & pID & " ;"
        ConSys.Execute stSQL1: stSQL5 = ""
        ANS = MsgBox("Deleted", vbInformation + vbOKOnly)
        Call f_List_Price(contractID)
        
        '一番若い料金期間を持つデータの、Lock を 外す ==================================================
        Set rs2 = New ADODB.Recordset
        stSQL2 = "SELECT MAX([ID]) AS [MxID] FROM CCM.[" & CCMCOST1 & "] WHERE [CCM_ID] = " & contractID & " ;"
        Set rs2 = New ADODB.Recordset
        rs2.Open stSQL2, ConSys, adOpenDynamic, adLockOptimistic
                If Not rs2.EOF Then
                        If IsNull(rs2![MxID]) = False Then
                                stSQL3 = "UPDATE CCM.[" & CCMCOST1 & "] SET  [Locked] = 0 WHERE [ID] = " & rs2![MxID] & " "
                                ConSys.Execute stSQL3: stSQL3 = ""
                        End If
                End If
        rs2.Close:  Set rs2 = Nothing:  stSQL2 = ""
        Exit Function
        
ERR_EXIT:
        Call f_List_Price(contractID)
End Function

Private Function pf_Check_Locked(ck As Boolean)
        Dim i As Integer
        With Forms(FRM03)
                If contractStatus = "Cancelled" Then
                        'Cancell の場合？？
                Else
                        If ck = True Then
                                For i = 3 To 15
                                        .Controls("Text" & 200 + i).Locked = True
                                Next i
                        ElseIf ck = False Then
                                For i = 3 To 15
                                        .Controls("Text" & 200 + i).Locked = False
                                Next i
                        End If
                End If
        End With
End Function

Public Function f_Check_Term(sDate As Date, eDate As Date, ud As Boolean) As Boolean
        f_Check_Term = False
        If sDate > eDate Then 'Textbox 内のDateに矛盾がないかチェック。
                ANS = MsgBox("End date must be later than Start date.", vbCritical + vbOKOnly)
                f_Check_Term = False:   Exit Function
        Else
                f_Check_Term = True
        End If
End Function

Public Function f_Check_Term_EndDate(sDate As Date, eDate As Date, ud As Boolean) As Boolean
        '料金は、遡って入力する事は出来ない。そのためのチェック。
        'Update の場合は、2番目に大きいEnd Date以降で、Update出来る。Update出来るのは、最新の料金データだけだから・
        'Newの場合は、一番大きいEnd Data以降でのみ、Updateできる。
        f_Check_Term_EndDate = False
        Dim stSQL1 As String, stSQL2 As String
        Dim rs1 As ADODB.Recordset, rs2 As ADODB.Recordset
        
        Set rs1 = New ADODB.Recordset
        Set rs2 = New ADODB.Recordset
                
        If ud = True Then   'When [Update], Get 2nd Largest End Date.
                stSQL1 = "SELECT COUNT([ID]) AS [pNum] FROM CCM.[" & CCMCOST1 & "] WHERE [CCM_ID] = " & contractID & ";"
                rs1.Open stSQL1, ConSys, adOpenDynamic, adLockReadOnly
                        Select Case rs1![pNum]
                                Case Is <= 1  '既存レコードが1個以下の時、無条件でUpdate可能に。
                                        stSQL2 = ""
                                Case Is >= 2:   '既存レコードが1個以上の時、２番目に大きいEnd Date でチェック。今Updateしようとしているのが、一番大きいDataだから。
                                        stSQL2 = "SELECT TOP 1 [term_end] AS [MaxEnd] FROM " & _
                                                        "(SELECT TOP 2 * FROM CCM.[" & CCMCOST1 & "] WHERE [CCM_ID] = " & contractID & " ORDER BY [term_end] DESC) AS A " & _
                                                        "ORDER BY [term_end] ASC"
                        End Select
                rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
        Else    'When [New], Get 1st Largest End Date (Newの時は最大のEnd Date)
                If contractID = 0 Then
                        f_Check_Term_EndDate = True
                        Exit Function
                Else
                        stSQL2 = "SELECT MAX([term_end]) AS [MaxEnd] FROM CCM.[" & CCMCOST1 & "] " & _
                                        "WHERE [CCM_ID] = " & contractID & ";"
                End If
        End If
        
        If Not stSQL2 = "" Then
                Debug.Print stSQL2
                rs2.Open stSQL2, ConSys, adOpenDynamic, adLockReadOnly
                If Not rs2.EOF Then
                        If Format(eDate, "yyyy/mm/dd") > Format(rs2![MaxEnd], "yyyy/mm/dd") Then
                                f_Check_Term_EndDate = True
                        End If
                        If Format(eDate, "yyyy/mm/dd") <= Format(rs2![MaxEnd], "yyyy/mm/dd") Then
                                ANS = MsgBox("End date must be after " & Format(rs2![MaxEnd], "yyyy/mm/dd"), vbExclamation + vbOKOnly)
                                Forms(FRM03).Text204 = ""
                                f_Check_Term_EndDate = False
                        End If
                End If
                rs2.Close:  Set rs2 = Nothing:  stSQL2 = ""
        Else
                f_Check_Term_EndDate = True
        End If
End Function

Public Function f_Check_Term_StartDate(sDate As Date, eDate As Date, ud As Boolean) As Boolean
        '料金は、遡って入力する事は出来ない。そのためのチェック。
        'Update の場合は、2番目に大きいEnd Date以降で、Update出来る。Update出来るのは、最新の料金データだけだから・
        'Newの場合は、一番大きいEnd Data以降でのみ、Updateできる。
        f_Check_Term_StartDate = False
        Dim stSQL1 As String, stSQL2 As String
        Dim rs1 As ADODB.Recordset, rs2 As ADODB.Recordset
        
        Set rs1 = New ADODB.Recordset
        Set rs2 = New ADODB.Recordset
        
        If ud = True Then   'When [Update], Get 2nd Largest End Date.
                stSQL1 = "SELECT COUNT([ID]) AS [pNum] FROM CCM.[" & CCMCOST1 & "] WHERE [CCM_ID] = " & contractID & ";"
                rs1.Open stSQL1, ConSys, adOpenDynamic, adLockReadOnly
                        Debug.Print rs1![pNum]
                        Select Case rs1![pNum]
                                Case Is <= 1  '既存レコードが1個以下の時、無条件でUpdate可能に。
                                        stSQL2 = ""
                                Case Is >= 2:   '既存レコードが1個以上の時、２番目に大きいEnd Date でチェック。今Updateしようとしているのが、一番大きいDataだから。
                                        stSQL2 = "SELECT TOP 1 [term_end] AS [MaxEnd] FROM " & _
                                                        "(SELECT TOP 2 * FROM CCM.[" & CCMCOST1 & "] WHERE [CCM_ID] = " & contractID & " ORDER BY [term_end] DESC) AS A " & _
                                                        "ORDER BY [term_end] ASC"
                        End Select
                rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
        Else    'When [New], Get 1st Largest End Date (Newの時は最大のEnd Date)
                If contractID = 0 Then
                        f_Check_Term_StartDate = True
                        Exit Function
                Else
                        stSQL2 = "SELECT MAX([term_end]) AS [MaxEnd] FROM CCM.[" & CCMCOST1 & "] " & _
                                        "WHERE [CCM_ID] = " & contractID & ";"
                End If
        End If
        
        If Not stSQL2 = "" Then
                Debug.Print stSQL2
                rs2.Open stSQL2, ConSys, adOpenDynamic, adLockReadOnly
                If Not rs2.EOF Then
                        If Format(sDate, "yyyy/mm/dd") > Format(rs2![MaxEnd], "yyyy/mm/dd") Then
                                f_Check_Term_StartDate = True
                        End If
                        If Format(sDate, "yyyy/mm/dd") <= Format(rs2![MaxEnd], "yyyy/mm/dd") Then
                                ANS = MsgBox("Start date must be after " & Format(rs2![MaxEnd], "yyyy/mm/dd"), vbExclamation + vbOKOnly)
                                Forms(FRM03).Text203 = ""
                                f_Check_Term_StartDate = False
                        End If
                End If
                rs2.Close:  Set rs2 = Nothing:  stSQL2 = ""
        Else
                f_Check_Term_StartDate = True
        End If
End Function

Public Function f_Calculate_Term(sDate As Date, eDate As Date) As Integer
        If Day(sDate) > Day(eDate) Then
                f_Calculate_Term = DateDiff("m", sDate, eDate)
        Else
                f_Calculate_Term = DateDiff("m", sDate, eDate) + 1
        End If
End Function

Public Function f_Monthly_Cost(pSchedule As String, pCost As Double, pTerm As Integer) As Double
        Dim stSQL1 As String, rs1 As ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        stSQL1 = "SELECT * FROM CCM.[" & MST04 & "] WHERE [Name] = '" & pSchedule & "';"
        rs1.Open stSQL1, ConSys, adOpenDynamic, adLockReadOnly
                If Not rs1.EOF Then
                        If rs1![divisor] = 0 Then
                                f_Monthly_Cost = pCost / pTerm
                        Else
                                f_Monthly_Cost = pCost / rs1![divisor]
                        End If
                End If
        rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
End Function

Public Function f_Monthly_Cost_Mod(pSchedule As String, pCost As Double, pTerm As Integer) As Double
        Dim stSQL1 As String, rs1 As ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        stSQL1 = "SELECT * FROM CCM.[" & MST04 & "] WHERE [Name] = '" & pSchedule & "';"
        rs1.Open stSQL1, ConSys, adOpenDynamic, adLockReadOnly
        If Not rs1.EOF Then
                If rs1![divisor] = 0 Then
                        f_Monthly_Cost_Mod = pCost Mod pTerm
                Else
                        f_Monthly_Cost_Mod = pCost Mod rs1![divisor]
                End If
        End If
        rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
End Function

Public Function f_Monthly_Cost_Fix(pSchedule As String, pCost As Double, pTerm As Integer) As Double
        Dim stSQL1 As String, rs1 As ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        stSQL1 = "SELECT * FROM CCM.[" & MST04 & "] WHERE [Name] = '" & pSchedule & "';"
        rs1.Open stSQL1, ConSys, adOpenDynamic, adLockReadOnly
        If Not rs1.EOF Then
                If rs1![divisor] = 0 Then
                        f_Monthly_Cost_Fix = Fix(pCost / pTerm)
                Else
                        f_Monthly_Cost_Fix = Fix(pCost / rs1![divisor])
                End If
        End If
        rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
End Function

Public Function f_Annually_Cost(pSchedule As String, pCost As Double, pTerm As Integer) As Double
        Dim stSQL1 As String, rs1 As ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        stSQL1 = "SELECT * FROM CCM.[" & MST04 & "] WHERE [Name] = '" & pSchedule & "';"
        rs1.Open stSQL1, ConSys, adOpenDynamic, adLockReadOnly
                If Not rs1.EOF Then
                        If rs1![divisor] = 0 Then
                                f_Annually_Cost = pCost * 12 / pTerm
                        Else
                                f_Annually_Cost = pCost * 12 / rs1![divisor]
                        End If
                End If
        rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
End Function

Public Function f_Total_Cost(pSchedule As String, pCost As Double, pTerm As Integer) As Double
        Dim stSQL1 As String, stSQL2 As String
        Dim rs1 As ADODB.Recordset, rs2 As ADODB.Recordset
        Dim tCost As Double
        Set rs1 = New ADODB.Recordset
        
        '変更した金額分の計算
        stSQL1 = "SELECT * FROM CCM.[" & MST04 & "] WHERE [Name] = '" & pSchedule & "';"
        rs1.Open stSQL1, ConSys, adOpenDynamic, adLockReadOnly
                If Not rs1.EOF Then
                        If rs1![divisor] = 0 Then
                                tCost = pCost
                        Else
                                tCost = pCost * pTerm / rs1![divisor]
                        End If
                End If
        rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
        
        '過去の金額の計算（リストに表示されている金額の、1番上を除く、全ての合計を出す）
        Dim i As Integer, j As Integer
        Set rs2 = New ADODB.Recordset
        With Forms(FRM03)
                If IsNumeric(.Text202) = True Then
                        j = 2
                Else
                        j = 1
                End If
                For i = j To .Text201.ListCount - 1
                        If Format(.Text201.Column(2, i), "yyyy/mm/dd") <> Format(.Text203, "yyyy/mm/dd") And Format(.Text201.Column(3, i), "yyyy,mm,dd") <> Format(.Text204, "yyyy/mm/dd") Then
                                stSQL2 = "SELECT * FROM CCM.[" & MST04 & "] WHERE [Name] = '" & .Text201.Column(6, i) & "';"
                                rs2.Open stSQL2, ConSys, adOpenDynamic, adLockReadOnly
                                        If Not rs2.EOF Then
                                                If rs2![divisor] = 0 Then
                                                        tCost = tCost + .Text201.Column(7, i)
                                                Else
                                                        tCost = tCost + .Text201.Column(7, i) * .Text201.Column(4, i) / rs2![divisor]
                                                End If
                                        End If
                                rs2.Close:  stSQL2 = ""
                        End If
                Next i
        End With
        Set rs2 = Nothing
        f_Total_Cost = tCost
        tCost = 0
End Function

Public Function f_Set_Term(strDate As Date, endDate As Date) As Integer
        f_Set_Term = 0
        Dim trm As Integer, adj As Integer
        trm = 0:    adj = 0
        
        If Forms(FRM01).Check01 = True Then
                adj = adj + 1
        End If
        If Forms(FRM01).Check02 = True Then
                adj = adj + 1
        End If
        
        trm = DateDiff("m", Format(strDate, "yyyy/mm/dd"), Format(endDate, "yyyy/mm/dd"))
         f_Set_Term = trm - 1 + adj
        trm = 0:    adj = 0
End Function

Public Function f_Monthly_Payment(ttlAmount As Double, divVal As Single) As Double
        If divVal = 0 Then
                f_Monthly_Payment = 0
        Else
                f_Monthly_Payment = Int(ttlAmount / divVal)
        End If
        ttlAmount = 0:  divVal = 0
End Function

Public Function f_Annual_Payment(ttlAmount As Double, divVal As Single, termVal As Single) As Double
        If divVal = 0 Then
                f_Annual_Payment = Int(ttlAmount / termVal * 12)
        Else
                f_Annual_Payment = Int(ttlAmount / divVal * 12)
        End If
        ttlAmount = 0:  divVal = 0
End Function

Public Function f_Payment_Schedule(strDate As Date, conMonths As Integer, amoMonth As Double, incl As Boolean)
        Dim stSQL1 As String
        Forms(FRM01).Form.Container01.Form.RecordSource = ""
        f_Delete_LocalTable (TMP01)
        stSQL1 = "CREATE TABLE " & TMP01 & "(ID COUNTER PRIMARY KEY, [Month] DATE, [Amount] DOUBLE)"
        DoCmd.RunSQL stSQL1
        
        Dim stSQL2 As String
        Dim db As DAO.Database
        Dim nMonth As Date
         
        Dim i As Integer
        For i = 0 To conMonths - 1
                nMonth = DateAdd("M", i + IIf(incl = True, 0, 1), strDate)
                stSQL2 = "INSERT INTO " & TMP01 & "([Month], [Amount]) VALUES(#" & Format(nMonth, "yyyy/mm/dd") & "#, " & amoMonth & " )"
                DoCmd.SetWarnings False
                DoCmd.RunSQL stSQL2
                DoCmd.SetWarnings True
        Next i
        
        Forms(FRM01).Form.Container01.Form.RecordSource = "SELECT * FROM " & TMP01 & ";"
        Forms(FRM01).Form.Container01.Form.Requery
End Function

Public Function f_Total_Amount_Check(ttlAmount As Double, period As Double)
        Dim stSQL1 As String, stSQL2 As String
        Dim rs1 As DAO.Recordset, rs2 As DAO.Recordset
        Dim rc1 As Integer, rc2 As Integer
        Dim i As Integer, j As Integer, difttl As Double
        Dim db As DAO.Database
        Set db = CurrentDb
        
        If period = 0 Then
                stSQL1 = "SELECT TOP 1 * FROM " & TMP01 & " ORDER BY ID"
                Set rs1 = db.OpenRecordset(stSQL1, dbOpenDynaset, dbSeeChanges)
                        rs1.Edit
                        rs1!amount = ttlAmount
                        rs1.Update
                rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
        Else
                stSQL1 = "SELECT SUM(Q1.Amount) AS Total FROM (SELECT TOP " & period & " * FROM " & TMP01 & " ORDER BY ID) AS Q1"
                Set rs1 = db.OpenRecordset(stSQL1, dbReadOnly)
                'Debug.Print rs1!Total & " / " & ttlAmount
                
                If rs1!Total <> ttlAmount Then

                        stSQL2 = "SELECT * FROM " & TMP01 & " ORDER BY ID"
                        Set rs2 = db.OpenRecordset(stSQL2, dbOpenDynaset, dbSeeChanges)
                                
                                rs2.MoveLast
                                rc2 = rs2.RecordCount
                                rs2.MoveFirst
                                i = 1
                                Debug.Print rc2 & " / " & period
                                If rc2 >= period Then
                                        difttl = ttlAmount - rs1!Total
                                        Debug.Print difttl & " = " & ttlAmount & " - " & rs1!Total
                        
                                        Do Until rs2.EOF
                                                If i <= period Then
                                                        If i <= difttl Then
                                                                rs2.Edit
                                                                rs2!amount = rs2!amount + 1
                                                                rs2.Update
                                                        End If
                                                        If i = period Then
                                                                i = 1
                                                        Else
                                                                i = i + 1
                                                        End If
                                                End If
                                                rs2.MoveNext
                                        Loop
                                Else
                                        difttl = (ttlAmount / period) * rc2 - rs1!Total
                                        Debug.Print difttl & " = (" & ttlAmount & " / " & period & ") * " & rc2 & " - " & rs1!Total
                                        
                                        Do Until rs2.EOF
                                                If i <= difttl Then
                                                    rs2.Edit
                                                    rs2!amount = rs2!amount + 1
                                                    rs2.Update
                                                    i = i + 1
                                                End If
                                                rs2.MoveNext
                                        Loop
                                        
                                End If
                End If
        End If
        Forms(FRM01).Form.Container01.Form.Requery
End Function

Private Function pf_Update_Monthly_Cost() As Boolean
        pf_Update_Monthly_Cost = False
        On Error GoTo ERR_MSG
        Dim cID As Long, ccmID As Long, ccmNum As String
        Dim sDate As Double, eDate As Double, rDate As Double, mTerm As Integer, pSchedule As String
        Dim pAmount As Double, mAmount As Double
        Dim cNonFix As Double, cMod As Long, cFix As Long
        Dim stSQL1 As String, stSQL2 As String
        Dim m As Integer
        
        'Monthly Cost を整数で入力するか少数で入力するかの切替 =--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--
        Dim i As Boolean
        i = True '(True: "整数"、False: "少数")
        
        With Forms(FRM03)
                '既存レコードの変更の場合、既存の価格を削除 ===================================================
                If IsNumeric(.Text202) = True Then
                        cID = .Text202
                        stSQL1 = "DELETE FROM CCM.[RPT_Monthly_Cost] WHERE [cost_ID] = " & cID & ";"
                        ConSys.Execute stSQL1: stSQL1 = ""
                End If
                
                '値の取得 ---------------
                sDate = .Text203
                mTerm = .Text205
                If contractStatus = "Cancelled" Then
                    eDate = .Text107
                    mTerm = DateDiff("m", DateSerial(Year(sDate), Month(sDate), 1), DateSerial(Year(eDate), Month(eDate), 1)) + 1
                End If
                pSchedule = .Text206
                pAmount = .Text208
                ccmID = .Text212
                ccmNum = .Text213
                mAmount = .Text210
                If i = True Then
                    cFix = f_Monthly_Cost_Fix(pSchedule, pAmount, mTerm)   '整数値のMonthly Costを算出
                    cMod = f_Monthly_Cost_Mod(pSchedule, pAmount, mTerm)   'Monthly Costの余りを算出
                Else
                    cNonFix = f_Monthly_Cost(pSchedule, pAmount, mTerm)
                End If
                
                '期間中全ての月の料金を、個別に登録 =========================================================
                For m = 0 To mTerm - 1
                        rDate = DateAdd("m", m, sDate)
                        rDate = DateSerial(Year(rDate), Month(rDate), 1)
                        
                        If i = True Then 'Monthly Costを整数で入力する場合 =--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--
                                'Monthly Costを整数値で入れる場合（あまり分を期間の初めにまぶす） ----------------------------------------------------
                                If m < cMod Then
                                        stSQL2 = "INSERT INTO CCM.[RPT_Monthly_Cost]([cost_ID], [CCM_ID], [CCM_number], [month], [cost]) " & _
                                                "VALUES(" & cID & ", " & ccmID & ", '" & ccmNum & "', " & rDate & ", " & cFix + 1 & ")"
                                
                                Else    'その他の月は、商の整数値のみ
                                        stSQL2 = "INSERT INTO CCM.[RPT_Monthly_Cost]([cost_ID], [CCM_ID], [CCM_number], [month], [cost]) " & _
                                                "VALUES(" & cID & ", " & ccmID & ", '" & ccmNum & "', " & rDate & ", " & cFix & ")"
                                End If
                        Else    'Monthly Cost を少数で入力する場合 '=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--
                                '小数点のまま入れる場合 -----------------------------------------------------------------------------------------------------------------------------
                                stSQL2 = "INSERT INTO CCM.[RPT_Monthly_Cost]([cost_ID], [CCM_ID], [CCM_number], [month], [cost]) " & _
                                                "VALUES(" & cID & ", " & ccmID & ", '" & ccmNum & "', " & rDate & ", " & cNonFix & ")"
                        End If
                        
                        ConSys.Execute stSQL2: stSQL2 = ""
                Next m
                cID = 0:    sDate = 0:  rDate = 0:  mTerm = 0: m = 0:   pSchedule = ""
                pAmount = 0: ccmID = 0:    ccmNum = "":    mAmount = 0
                cMod = 0:   cFix = 0: cNonFix = 0
        End With
        
        pf_Update_Monthly_Cost = True
        Exit Function
        
ERR_MSG: 'エラー処理 -----------
        pf_Update_Monthly_Cost = False
        ANS = MsgBox("Error on Saving Monthly Data for Variance report." & vbCrLf & "Please ask system adminstrator.", vbInformation + vbOKOnly)
        Exit Function
End Function

Public Function f_Cancel_Monthly_Cost(sDate As Double, eDate As Double) As Boolean
        f_Cancel_Monthly_Cost = False
        Dim rDate As Double, stSQL1 As String
        rDate = DateAdd("m", f_Calculate_Term(Format(sDate, "yyyy/mm/dd"), Format(eDate, "yyyy/mm/dd")), DateSerial(Year(sDate), Month(sDate), 1))
        Debug.Print "m: " & f_Calculate_Term(Format(sDate, "yyyy/mm/dd"), Format(eDate, "yyyy/mm/dd")) & " / Last: " & Format(rDate, "yyyy/mm/dd")
        If contractStatus = "Cancelled" Then
                stSQL1 = "DELETE FROM CCM.[RPT_Monthly_Cost] WHERE [CCM_number] = '" & contractNumber & "' AND [month] > " & rDate & ";"
                ConSys.Execute stSQL1
        End If
        f_Cancel_Monthly_Cost = True
End Function
