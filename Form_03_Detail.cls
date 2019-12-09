VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_03_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database:    Option Explicit

Private Sub Form_Load()
        Call pf_Authentication_FRM03
        Call pf_Button_Reset
        Me.Text118.RowSource = "Reccuring;Onetime"
        If contractID = 0 And contractNumber = "" Then
                Call ComContractNew_Click
        Else
                Call pf_Reset_Contract_Info
        End If
End Sub

Private Sub Form_Close()
        Call pf_Clear_Combobox
        DoCmd.OpenForm FRM01
End Sub

Private Sub ComClose_Click()
        Call pf_Clear_Combobox
        DoCmd.Close acForm, FRM03
End Sub

'===== CONTRACT ====================================================================================================
'===================================================================================================================
Private Sub ComContractEdit_Click()
        If pf_Button_Cancel(ComContractEdit) = False Then
                Select Case f_Button_Unlock_Item(contractStatus)
                        Case 0
                                Call pf_Button_Enable_Contract(1, 1, 0, 0)
                                Call pf_Button_Enable_Price(0, 0, 0, 0)
                                Call pf_Textbox_Unlock_Item(authLevel, False)
                                Me.Text112.Locked = True
                                Me.Text112.BackColor = RGB(236, 236, 236)
                                Me.Text201.Locked = True
                        Case 1
                                Call pf_Button_Enable_Contract(1, 0, 1, 1)
                End Select
        Else
                Call pf_Button_Reset
                Call pf_Reset_Contract_Info
                Call f_Check_Item_Cancel(contractStatus)
                Me.Text201.Locked = False
        End If
End Sub

Private Sub ComContractNew_Click()
        If pf_Button_Cancel(ComContractNew) = False Then
                Call pf_ID_Store(1)
                Call pf_Button_Enable_Contract(0, 1, 1, 0)
                Call pf_Textbox_Unlock_Item(authLevel, True)
                Call pf_Button_Enable_Price(0, 0, 0, 0)
                Call pf_Textbox_Unlock_Price(authLevel, True)
                Call pf_Button_Enable_Attachment(0, 0, 0)
                Call pf_TextBox_For_NewContract(0)
                Me.Text201.Locked = True
                'Call f_Button_Unlock_Item("Active")
        Else
                If tmpContractID = 0 Then
                        DoCmd.Close acForm, FRM03
                Else
                        Call pf_ID_Store(0)
                        Call pf_Button_Reset
                        Call pf_Reset_Contract_Info
                        Call f_Check_Item_Cancel(contractStatus)
                        Me.Text201.Locked = False
                End If
        End If
End Sub

Private Sub ComContractCopy_Click()
        If pf_Button_Cancel(ComContractCopy) = False Then
                Call pf_ID_Store(1)
                Call pf_Button_Enable_Contract(0, 1, 0, 1)
                Call pf_Textbox_Unlock_Item(authLevel, False)
                Call pf_Button_Enable_Price(0, 0, 0, 0)
                Call pf_Textbox_Unlock_Price(authLevel, True)
                Call pf_Button_Enable_Attachment(0, 0, 0)
                Call pf_TextBox_For_NewContract(1)
                Me.Text201.Locked = True
                'Call f_Button_Unlock_Item("Active")
        Else
                If tmpContractID = 0 Then
                        DoCmd.Close acForm, FRM03
                Else
                        Call pf_ID_Store(0)
                        Call pf_Button_Reset
                        Call pf_Reset_Contract_Info
                        Call f_Check_Item_Cancel(contractStatus)
                        Me.Text201.Locked = False
                End If
        End If
End Sub

Private Sub ComContractUpdate_Click()
        If IsNumeric(Me.Text101) = True Then    'ContractのIDがある場合、既存のContractになるので、Updateとなる。 ===================================================
                Me.Text131 = userName:  Me.Text132 = Now()
                
                Select Case f_Check_Change_Status(Me.Text102, Me.Text103)
                        Case 0  'ステータスが逆行する場合は、Update出来ない。
                                ANS = MsgBox("Backwords status change is not allowed. Please create THE NEW CONTRACT.", vbExclamation + vbOKOnly, "Disallowed status change")
                                
                        Case 1  'ステータスの変更が無くて、もともとのステータスがCancel属性の場合は、Updateできない。
                                ANS = MsgBox("Cancelled contract cannot be modified. Please create A NEW CONTRACT.", vbExclamation + vbOKOnly, "Disallowed status change")
                                Call pf_Reset_Contract_Info
                                Call pf_Button_Reset
                                Call f_Check_Item_Cancel(contractStatus)
                                
                        Case 2  'ステータスの変更がなくて、もともとのステータスがActive属性の場合は、Updateする。
                                If f_Update_Check(Me.Text103) = True Then
                                        ANS = MsgBox("Are you sure to UPDATE?", vbQuestion + vbYesNo, "Update Confirmation")
                                        If ANS = vbYes Then
                                                If pf_Update = True Then
                                                        ANS = MsgBox("Updated.", vbInformation + vbOKOnly, "Update Confirmation")
                                                End If
                                        End If
                                Else
                                        ANS = MsgBox("Please fill mandatory fields.", vbExclamation + vbOKOnly, "Mandatory Confirmation")
                                End If
                                
                        Case 3  'ステータスがCancel属性に変更される場合は、Cancel Dateの入力を確認する。
                                If f_Update_Check(Me.Text103) = True Then
                                        ANS = MsgBox("Are you sure to UPDATE?", vbQuestion + vbYesNo, "Update Confirmation")
                                        If ANS = vbYes Then
                                                contractStatus = Me.Text103
                                                If pf_Update = True Then
                                                        ANS = MsgBox("Updated.", vbInformation + vbOKOnly, "Update Confirmation")
                                                End If
                                        End If
                                Else
                                        ANS = MsgBox("Please fill mandatory fields.", vbExclamation + vbOKOnly, "Mandatory Confirmation")
                                End If
                                
                        Case 4  'ステータス変更が正常の場合で、新しいステータスがActive属性の場合、Updateする。
                                contractStatus = Me.Text103
                                Debug.Print Me.Text201.ListCount
                                'If Nz(Me.Text201, "") <> "" Then     'Priceの登録が既にあったら (もともと 「If Me.Text201.ListCount > 0 Then」だったが、何故か何回か遷移してくると、Text201のListCount が1になってしまう =============
                                If Me.Text201.ListCount <> 0 Then
                                        If f_Update_Check(Me.Text103) = True Then
                                                ANS = MsgBox("Are you sure to UPDATE?", vbQuestion + vbYesNo, "Update Confirmation")
                                                If ANS = vbYes Then
                                                        If pf_Update = True Then
                                                                ANS = MsgBox("Updated.", vbInformation + vbOKOnly, "Update Confirmation")
                                                        End If
                                                End If
                                        Else
                                                ANS = MsgBox("Please fill mandatory fields.", vbExclamation + vbOKOnly, "Mandatory Confirmation")
                                        End If
                                Else        'Priceの登録がなかったら ==================================
                                        If f_Update_Check(Me.Text103) = True And f_Update_Check_Price(Me.Text103) = True Then
                                                ANS = MsgBox("Are you sure to UPDATE?", vbQuestion + vbYesNo, "Update Confirmation")
                                                If ANS = vbYes Then
                                                        If pf_Update = True Then
                                                                If f_Get_New_ID(Me.Text102) = False Then Exit Sub
                                                                Me.Text212 = contractID
                                                                Me.Text213 = contractNumber
                                                                Call f_New_Price
                                                                ANS = MsgBox("Updated.", vbInformation + vbOKOnly, "Update Confirmation")
                                                        End If
                                                End If
                                        Else
                                                ANS = MsgBox("Please fill mandatory fields.", vbExclamation + vbOKOnly, "Mandatory Confirmation")
                                        End If
                                End If
                End Select
                
        Else    'Contract の IDが無い場合、新しいContractの入力となる。=====================================================================================
                ANS = MsgBox("Are you sure to create the NEW RECORD?", vbQuestion + vbYesNo, "New Create Confirmation")
                If ANS = vbYes Then
                        Call pf_New
                End If
        End If
        Me.Text201.Locked = False
End Sub

Private Function pf_Update() As Boolean
        pf_Update = False
        On Error GoTo ERR_EXIT
        
        Me.Text131 = userName:  Me.Text132 = Now()  'Updated PIC/時間 の強制入力 ==========================================
        If f_Update_CCM = True Then
                If IsNumeric(Me.Text141) = True Then
                        Call f_Update_BOID(Me.Text141)
                End If
                If contractStatus = "Cancelled" Then
                        Call f_Cancel_Monthly_Cost(Me.Text105, Me.Text107)
                End If
                pf_Update = True
        Else
ERR_EXIT:
                ANS = MsgBox("Something went wrong, the data might have error." & vbCrLf & _
                                        "Please ask administrator to check the data.", vbCritical + vbOKOnly)
                pf_Update = False
        End If
        
        Call f_Cancelling
        Call pf_Button_Reset
        Call pf_Reset_Contract_Info
        
End Function

Private Function pf_New() As Boolean
        pf_New = False
        On Error GoTo ERR_EXIT
        Me.Text131 = userName:  Me.Text132 = Now()  'Updated PIC/時間 の強制入力 ==========================================
        
        Select Case Me.Text103
                Case "Active"
                        If f_Update_Check(Me.Text103) = True And f_Update_Check_Price(Me.Text103) = True Then
                                'Contract 情報のInsert --------------------------------------------------------------------------------------------------
                                If f_New_CCM = False Then
                                        ANS = MsgBox("Something went wrong on the Property Data, the data might have error." & vbCrLf & _
                                                                "Please ask administrator to check the data.", vbCritical + vbOKOnly)
                                        GoTo ERR_EXIT
                                End If

                                'BOID情報のInsert -------------------------------------------------------------------------------------------------------
                                If IsNumeric(Me.Text141) = True Then
                                        If f_New_BOID(Me.Text141) = False Then
                                                ANS = MsgBox("Something went wrong on the BOID Data, the data might have error." & vbCrLf & _
                                                                        "Please ask administrator to check the data.", vbCritical + vbOKOnly)
                                                GoTo ERR_EXIT
                                        End If
                                End If
                                
                                'Contract/BOID 情報の再表示 ---------------------------------------------------------------------------------------
                                If f_Get_New_ID(Me.Text102) = False Then Exit Function
                                Call f_Choose_Item(contractID)
                                Me.Text212 = contractID:    Me.Text213 = contractNumber
                                
                                'Price情報のInsert --------------------------------------------------------------------------------------------------------
                                If IsDate(Me.Text203) = True And IsDate(Me.Text204) = True Then
                                        If f_New_Price = False Then
                                                ANS = MsgBox("Something went wrong on the Price Data, the data might have error." & vbCrLf & _
                                                                        "Please ask administrator to check the data.", vbCritical + vbOKOnly)
                                                GoTo ERR_EXIT
                                        End If
                                End If

                                'Insert終了 / Serch画面のリフレッシュ --------------------------------------------------------------------------
                                Call pf_Button_Reset
                                Call f_List_Contract(fieldCount)
                                ANS = MsgBox("Inserted the New contract.", vbInformation + vbOKOnly)
                                Forms(FRM01).SUB001.Form.Requery
                                pf_New = True
                        Else
                                'Mandatory 不足 ----------------------------------------------------------------------------------------------------------
                                ANS = MsgBox("Please fill mandatory fields.", vbExclamation + vbOKOnly)
                                pf_New = False
                        End If

                Case "Draft"
                        If f_Update_Check(Me.Text103) = True Then
                                'Contract 情報のInsert --------------------------------------------------------------------------------------------------
                                If f_New_CCM = False Then
                                        ANS = MsgBox("Something went wrong on the Property Data, the data might have error." & vbCrLf & _
                                                                "Please ask administrator to check the data.", vbCritical + vbOKOnly)
                                        GoTo ERR_EXIT
                                End If
                                
                                If f_Get_New_ID(Me.Text102) = False Then Exit Function
                                
                                'BOID情報のInsert -------------------------------------------------------------------------------------------------------
                                If IsNumeric(Me.Text141) = True Then
                                        If f_New_BOID(Me.Text141) = False Then
                                                ANS = MsgBox("Something went wrong on the BOID Data, the data might have error." & vbCrLf & _
                                                                        "Please ask administrator to check the data.", vbCritical + vbOKOnly)
                                                GoTo ERR_EXIT
                                        End If
                                End If
                                
                                'Contract/BOID 情報の再表示 ---------------------------------------------------------------------------------------
                                Call f_Choose_Item(contractID)
                                
                                'Price情報のInsert --------------------------------------------------------------------------------------------------------
                                If IsDate(Me.Text203) = True And IsDate(Me.Text204) = True Then
                                        Me.Text212 = contractID:    Me.Text213 = contractNumber
                                        If f_New_Price = False Then
                                                ANS = MsgBox("Something went wrong on the Price Data, the data might have error." & vbCrLf & _
                                                                        "Please ask administrator to check the data.", vbCritical + vbOKOnly)
                                                GoTo ERR_EXIT
                                        End If
                                Else
                                        Me.Text212 = "":    Me.Text213 = ""
                                End If
                                
                                'Insert終了 / Serch画面のリフレッシュ --------------------------------------------------------------------------
                                Call pf_Button_Reset
                                Call f_List_Contract(fieldCount)
                                ANS = MsgBox("Inserted the New contract.", vbInformation + vbOKOnly)
                                Forms(FRM01).SUB001.Form.Requery
                                pf_New = True
                                
                        Else
                                'Mandatory 不足 ----------------------------------------------------------------------------------------------------------
                                ANS = MsgBox("Please fill mandatory fields.", vbExclamation + vbOKOnly)
                                pf_New = False
                        End If
                        
                Case Else
                        'NG: いきなりCancel Status -------------------------------------------------------------------------------------------------
                        Call pf_Button_Reset
                        Call pf_Reset_Contract_Info
                        Call f_Check_Item_Cancel(contractStatus)
                        ANS = MsgBox("You cannot create 'Cancelled' Contract.", vbCritical + vbOKOnly)
                        pf_New = False
        End Select
        Call pf_Textbox_Lock_Price
        Exit Function
ERR_EXIT:
        'On Error 何も変更しない ---------------------------------------------------------------------------------------------------------------------
        Call pf_Button_Reset
        pf_New = False
End Function

Private Sub Text103_Change()
        If contractStatus = "" Then
                Call pf_Textbox_Unlock_Price(authLevel, False)
        ElseIf Me.Text103 = "Active" And contractStatus = "Draft" Then
                Call pf_Textbox_Unlock_Price(authLevel, False)
        Else
                Call pf_Textbox_Lock_Price
        End If
        Debug.Print "STTS:" & contractStatus
End Sub

Private Sub Text110_AfterUpdate()   'Address
        Me.Text111 = ""
        Me.Text111.RowSource = "SELECT [ID], [Pref_ID], [Address_E] FROM " & ConSQL & ".[CCM." & MST13 & "] WHERE [Pref_ID] = " & Me.Text110.Column(0) & ";"
End Sub

Private Sub Text112_Click()
        Me.Text113 = ""
        Me.Text113.RowSource = "SELECT DISTINCT [contract_type] FROM " & ConSQL & ".[CCM." & MST06 & "] WHERE [contract_model] = '" & Me.Text112 & "';"
        If IsNumeric(Me.Text101) = False Then
                Call f_Get_MaxNumber(Me.Text112)
                Me.Text102 = contractNumber
                Me.Text212 = contractNumber
        End If
        Call f_Show_Extra(Nz(Me.Text112, "Error"), Nz(Me.Text113, "Error"))
End Sub

Private Sub Text113_Click() 'Contract Type が決まったら、extra Fieldを表示する。
        Call f_Show_Extra(Nz(Me.Text112, "Error"), Nz(Me.Text113, "Error"))
End Sub

Private Sub Text120_Click() 'Initiative ID
        Dim stSQL1 As String
        Dim rs1 As DAO.Recordset
        stSQL1 = "SELECT * FROM " & ConSQL & ".[CCM." & MST08 & "] WHERE [Initiative_ID] = '" & Me.Text120 & "';"
        Set rs1 = db.OpenRecordset(stSQL1, dbReadOnly)
        If Not rs1.EOF Then
                Me.Text121 = rs1![Initiative_Name] 'Initiative Name
        End If
        rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
End Sub

Private Sub Text123_Click() 'Service Category
        Dim stSQL1 As String
        Dim rs1 As DAO.Recordset
        stSQL1 = "SELECT * FROM " & ConSQL & ".[CCM." & MST05 & "] WHERE [service_category] = '" & Me.Text123 & "';"
        Set rs1 = db.OpenRecordset(stSQL1, dbReadOnly)
        If Not rs1.EOF Then
                Me.Text122 = rs1![Department]   'Department
                Me.Text124 = rs1![payment_share_code] 'Payment Share Code
                Me.Text125 = rs1![CBD_Name]  'CBD Code
                Me.Text126 = rs1![mother_service] 'Mother Service
        End If
        rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
End Sub

Private Sub Text141_Change()
        If IsNumeric(Me.Text141) = True Then
                Me.Text142 = ""
                Me.Text142.RowSource = "SELECT DISTINCT [BO_ID] FROM " & ConSQL & ".[CCM." & MST16 & "] WHERE [year] = " & Me.Text141 & ";"
        End If
End Sub

Private Sub Text141_Click()
        If IsNumeric(Me.Text141) = True Then
                Me.Text142 = ""
                Me.Text142.RowSource = "SELECT DISTINCT [BO_ID] FROM " & ConSQL & ".[CCM." & MST16 & "] WHERE [year] = " & Me.Text141 & ";"
        End If
End Sub


'===== COST =========================================================================================================
'===================================================================================================================
Private Sub Text201_Click()
        priceID = Me.Text201
        Call f_Choose_Price(priceID)
        Call pf_Button_Reset
        Call f_Control_Price_Locked(0)
        Call f_Check_Item_Cancel(contractStatus)
        
        If authLevel >= 90 Then
                Call pf_Button_Enable_Price(1, 0, 1, 0)
                Me.ComPriceDel.Visible = True:  Me.ComPriceDel.Enabled = True
        Else
                Me.ComPriceDel.Visible = False:  Me.ComPriceDel.Enabled = False
                If f_Check_Cancel_Price(contractStatus) = False Then
                        If Me.Text215 = False Then
                                Call pf_Button_Enable_Price(1, 0, 1, 0)
                        Else
                                Call pf_Button_Enable_Price(0, 0, 1, 0)
                        End If
                Else
                        Call pf_Button_Enable_Price(0, 0, 0, 0)
                End If
        End If
End Sub

Private Sub ComPriceEdit_Click()
        If Me.Text201.ListIndex > -1 And Me.Text215 = False Then
                If pf_Button_Cancel(ComPriceEdit) = False Then
                        Select Case f_Button_Unlock_Price(contractStatus)
                                Case 0
                                        Call pf_Button_Enable_Contract(0, 0, 0, 0)
                                        Call pf_Button_Enable_Price(1, 1, 0, 0)
                                        Call pf_Textbox_Unlock_Price(authLevel, False)
                                Case 1
                                        Call pf_Button_Enable_Price(1, 0, 1, 0)
                                        Call pf_Button_Reset
                        End Select
                Else
                        Call pf_Button_Reset
                        Call f_List_Price(contractID)
                        Call f_Choose_Price(priceID)
                        Call f_Check_Cancel_Price(contractStatus)
                        Me.Text201.Selected(Me.Text201.ListIndex + 1) = False
                End If
        End If
End Sub

Private Sub ComPriceNew_Click()
        If pf_Button_Cancel(ComPriceNew) = False Then
                Call pf_Button_Enable_Contract(0, 0, 0, 0)
                Call pf_Button_Enable_Price(0, 1, 1, 0)
                Call pf_Textbox_Unlock_Price(authLevel, True)
                Call pf_TextBox_For_NewPrice(0)
        Else
                If contractID = 0 Then
                Else
                        Call pf_Button_Reset
                        Call f_List_Price(contractID)
                        Call f_Choose_Price(priceID)
                        Call f_Check_Cancel_Price(contractStatus)
                End If
        End If
End Sub

Private Sub ComPriceCopy_Click()
        'No Use
End Sub

Private Sub ComPriceUpdate_Click()
        ANS = MsgBox("Are you sure to UPDATE?", vbQuestion + vbYesNo)
        If ANS = vbYes Then
                If f_Update_Check_Price(Me.Text103) = True Then   'When Mandatory fields are Filled.
                        Me.Text210 = f_Monthly_Cost(Me.Text206, Me.Text208, Me.Text205)
                        Me.Text211 = f_Annually_Cost(Me.Text206, Me.Text208, Me.Text205)
                        Me.Text209 = f_Total_Cost(Me.Text206, Me.Text208, Me.Text205)
                        Me.Text201.RowSource = "":  Me.Text201.Requery
                        If IsNumeric(Me.Text202) = True Then 'Update or New
                                If f_Update_Price(Me.Text202) = True Then
                                        ANS = MsgBox("Updated.", vbInformation + vbOKOnly)
                                End If
                        Else
                                If f_New_Price = True Then
                                        ANS = MsgBox("Inserted the New Price.", vbInformation + vbOKOnly)
                                End If
                        End If
                        Call f_Check_Cancel_Price(contractStatus)
                        Call pf_Button_Reset
                        Call f_List_Price(contractID)
                Else        'When Mandatory fields are NOT Filled
                        ANS = MsgBox("Please fill mandatory fields.", vbExclamation + vbOKOnly)
                        Exit Sub
                End If
                Me.Text201.Selected(Me.Text201.ListIndex + 1) = False
        Else
                Call f_List_Price(contractID)
                Call pf_Button_Reset
        End If
        Me.Text201.Selected(Me.Text201.ListIndex + 1) = False
End Sub

Private Sub ComPriceDel_Click()
        ANS = MsgBox("Do you want to Delete the record?", vbQuestion + vbYesNo)
        If ANS = vbYes Then
                Debug.Print Me.Text201
                Call f_Delete_Price_Data(Me.Text202)
        End If
End Sub

Private Sub Text203_AfterUpdate()
        Dim sDate As Date
        Dim eDate As Date
        If IsNumeric(Me.Text203) = True Or IsDate(Me.Text203) = True Then
                sDate = Me.Text203
                If f_Check_Term_StartDate(sDate, eDate, IsNumeric(Me.Text202)) = False Then
                        Me.Text203 = "":    Exit Sub
                End If
        End If
        If IsNumeric(Me.Text204) = True Or IsDate(Me.Text204) = True Then
                eDate = Me.Text204
                If f_Check_Term_EndDate(sDate, eDate, IsNumeric(Me.Text202)) = False Then
                        Me.Text204 = "":    Exit Sub
                End If
        End If
        If (IsNumeric(Me.Text203) = True Or IsDate(Me.Text203) = True) And (IsNumeric(Me.Text204) = True Or IsDate(Me.Text204) = True) Then
                sDate = Me.Text203
                eDate = Me.Text204
                If f_Check_Term(sDate, eDate, IsNumeric(Me.Text202)) = False Then
                        Me.Text204 = "":    Exit Sub
                End If
                Me.Text205 = f_Calculate_Term(sDate, eDate)
                Me.Text105 = sDate
                Me.Text106 = eDate
                Call Text206_Change
        End If
End Sub

Private Sub Text204_AfterUpdate()
        Dim sDate As Date
        Dim eDate As Date
        If IsNumeric(Me.Text203) = True Or IsDate(Me.Text203) = True Then
                sDate = Me.Text203
                If f_Check_Term_StartDate(sDate, eDate, IsNumeric(Me.Text202)) = False Then
                        Me.Text203 = "":    Exit Sub
                End If
        End If
        If IsNumeric(Me.Text204) = True Or IsDate(Me.Text204) = True Then
                eDate = Me.Text204
                If f_Check_Term_EndDate(sDate, eDate, IsNumeric(Me.Text202)) = False Then
                        Me.Text204 = "":    Exit Sub
                End If
        End If
        If (IsNumeric(Me.Text203) = True Or IsDate(Me.Text203) = True) And (IsNumeric(Me.Text204) = True Or IsDate(Me.Text204) = True) Then
                sDate = Me.Text203
                eDate = Me.Text204
                If f_Check_Term(sDate, eDate, IsNumeric(Me.Text202)) = False Then
                        Me.Text204 = "":    Exit Sub
                End If
                Me.Text205 = f_Calculate_Term(sDate, eDate)
                Me.Text105 = sDate
                Me.Text106 = eDate
                Call Text206_Change
        End If
End Sub

Private Sub Text206_Change()
        If Nz(Me.Text206, "") <> "" And IsNumeric(Me.Text208) = True And IsNumeric(Me.Text205) = True Then
                If Nz(Me.Text204, 0) > 0 Then
                        Me.Text210 = f_Monthly_Cost(Me.Text206, Me.Text208, Me.Text205)
                        Me.Text211 = f_Annually_Cost(Me.Text206, Me.Text208, Me.Text205)
                        Me.Text209 = f_Total_Cost(Me.Text206, Me.Text208, Me.Text205)
                End If
        End If
End Sub

Private Sub Text208_AfterUpdate()
        If Nz(Me.Text206, "") <> "" And IsNumeric(Me.Text208) = True And IsNumeric(Me.Text205) = True Then
                If Nz(Me.Text204, 0) > 0 Then
                        Me.Text210 = f_Monthly_Cost(Me.Text206, Me.Text208, Me.Text205)
                        Me.Text211 = f_Annually_Cost(Me.Text206, Me.Text208, Me.Text205)
                        Me.Text209 = f_Total_Cost(Me.Text206, Me.Text208, Me.Text205)
                End If
        End If
End Sub


'===== ATTACHMENT ==================================================================================================
'===================================================================================================================
Private Sub ComAttachAdd_Click()
        If f_File_PickUp("Select upload file", "*.*") = True Then
                Dim i As Integer
                If f_Attachment_Upload(contractID, filePath) = False Then
                        'ANS = MsgBox("Error", vbCritical + vbOKOnly)
                Else
                        ANS = MsgBox("Attachment Uploaded.", vbInformation + vbOKOnly)
                End If
        End If
        Me.Text301.Requery
End Sub

Private Sub ComAttachDelete_Click()
        If Nz(Me.Text301, "") <> "" Then
                ANS = MsgBox("Are you really sure you delete the Attachment?" & vbCrLf & _
                                        "This process delete the attachment completely.", vbCritical + vbYesNo)
                If ANS = vbYes Then
                        If f_Attachment_Delete(Val(Me.Text301)) = True Then
                                ANS = MsgBox("Attachment Completely Deleted.", vbInformation + vbOKOnly)
                        End If
                Else
                        ANS = MsgBox("Deletion aborted.", vbInformation + vbOKOnly)
                End If
                Me.Text301.Requery
        End If
        Me.Text301 = ""
End Sub

Private Sub ComAttachDownload_Click()
        If Not Nz(contractNumber, "") = "" Then
            Call f_Attachment_Download(contractNumber)
        End If
End Sub


'=== PRIVATE FUNCTION =========================================================================
'=============================================================================================
Private Function pf_Textbox_Unlock_Item(c As Integer, n As Boolean)
        'Textbox と ComboboxのLockedを変更する。Textbox, Combobox の TagにLevelを入力してあるので、そのLevelによってLockedを変更するしないの判断としている。
        'これにより、いろいろな条件でLockedをTrueにしたりFalseにしたり出来る。また、nがTrueにより、Controlの中身を空に出来る。
        On Error Resume Next
        Dim nTag As Integer:    Dim meControl As Control
        For Each meControl In Me.Controls
                With meControl
                        If meControl.Name Like "Text1*" Then
                                nTag = Nz(meControl.Properties("tag"), 0)
                                If n = True Then meControl = ""
                                If nTag < c Then
                                        meControl.Locked = False
                                        meControl.BackColor = RGB(255, 255, 255)
                                Else
                                        meControl.Locked = True
                                        meControl.BackColor = RGB(236, 236, 236)
                                End If
                        End If
                End With
        Next
        On Error GoTo 0
End Function

Private Function pf_Textbox_Unlock_Price(c As Integer, n As Boolean)
        On Error Resume Next
        Dim nTag As Integer:    Dim meControl As Control
        For Each meControl In Me.Controls
                With meControl
                        If meControl.Name Like "Text2*" Then
                                nTag = Nz(meControl.Properties("tag"), 0)
                                If n = True Then meControl = ""
                                If nTag < c Then
                                        meControl.Locked = False
                                        meControl.BackColor = RGB(255, 255, 255)
                                Else
                                        meControl.Locked = True
                                        meControl.BackColor = RGB(236, 236, 236)
                                End If
                        End If
                End With
        Next
        On Error GoTo 0
End Function

Private Function pf_Textbox_Lock_Price()
        On Error Resume Next
        Dim nTag As Integer:    Dim meControl As Control
        For Each meControl In Me.Controls
                With meControl
                        If meControl.Name Like "Text2*" And Not meControl.Name Like "Text201" Then
                                meControl.Locked = True
                                meControl.BackColor = RGB(236, 236, 236)
                        End If
                End With
        Next
        On Error GoTo 0
End Function

Private Function pf_TextBox_For_NewContract(clearType As Integer)
        Me.Text101 = "":    Me.Text102 = "":    Me.Text103 = "Active"
        Me.Text105 = "": Me.Text106 = "":   Me.Text133 = ""
        Me.Text113.RowSource = ""
        Me.Text131 = userName
        Me.Text201.RowSource = ""
        Me.Text202 = "":    Me.Text203 = "":    Me.Text204 = "":    Me.Text205 = 0
        Me.Text206 = "": Me.Text207 = "JPY"
        Me.Text208 = "": Me.Text209 = 0: Me.Text210 = 0: Me.Text211 = 0
        Me.Text212 = Me.Text101
        Me.Text213 = Me.Text102
        Me.Text215 = 0
        Me.Text301.RowSource = ""
        Select Case clearType
                Case 0  'When New
                        Me.Text120 = "":    Me.Text121 = ""
                        Me.Text122 = "":    Me.Text124 = "": Me.Text125 = "":   Me.Text126 = ""
                        Me.Text129 = "Serv99999"
                Case 1  'When Copy
                        Me.Text102 = f_Get_MaxNumber(Me.Text112)
        End Select
End Function

Private Function pf_TextBox_For_NewPrice(clearType As Integer)
        Me.Text202 = ""
        Me.Text203 = "":    Me.Text204 = "":    Me.Text205 = ""
        Me.Text206 = ""
        Me.Text207 = "JPY"
        Me.Text208 = "": Me.Text209 = 0: Me.Text210 = 0: Me.Text211 = 0
        Me.Text212 = Me.Text101
        Me.Text213 = Me.Text102
        Me.Text215 = 0
        Select Case clearType
                Case 0  'When New
                        'Do nothing
                Case 1  'When Copy
                        'Do nothing
        End Select
End Function

Private Sub Text103_GotFocus()
        If Me.Text103.RowSource = "" Then
        Me.Text103.RowSource = "SELECT [status] FROM " & ConSQL & ".[CCM." & MST03 & "] WHERE [use] = true;"
        End If
End Sub

Private Sub Text110_GotFocus()
        If Me.Text110.RowSource = "" Then
        Me.Text110.RowSource = "SELECT DISTINCT [Pref_ID], [Pref_Name_J], [Pref_Name_E] FROM " & ConSQL & ".[CCM." & MST13 & "] ORDER BY [Pref_ID];"
        End If
End Sub

Private Sub Text112_GotFocus()
        If Me.Text112.RowSource = "" Then
        Me.Text112.RowSource = "SELECT DISTINCT [contract_model] FROM " & ConSQL & ".[CCM." & MST06 & "] ORDER BY [contract_model];"
        End If
End Sub

Private Sub Text113_GotFocus()
        If Me.Text113.RowSource = "" And Me.Text113.RowSource = "" Then
        Me.Text113.RowSource = "SELECT DISTINCT [contract_type] FROM " & ConSQL & ".[CCM." & MST06 & "] WHERE [contract_model] = '" & Me.Text112 & "' ORDER BY [contract_type];"  'Contract Modelを変更したら、変更・表示させる
        End If
End Sub

Private Sub Text114_GotFocus()
        If Me.Text114.RowSource = "" Then
        Me.Text114.RowSource = "SELECT [User_Name] FROM " & ConSQL & ".[CCM." & MST01 & "] WHERE [contract_Administrator] = true;" 'User Name はDLしない為
        End If
End Sub

Private Sub Text115_GotFocus()
        If Me.Text115.RowSource = "" Then
        Me.Text115.RowSource = "SELECT [vendor] FROM " & ConSQL & ".[CCM." & MST07 & "] WHERE [use] = true;"
        End If
End Sub

Private Sub Text117_GotFocus()
        If Me.Text117.RowSource = "" Then
        Me.Text117.RowSource = "SELECT [number] FROM " & ConSQL & ".[CCM." & MST14 & "] ORDER BY [number];"
        End If
End Sub

Private Sub Text119_GotFocus()
        If Me.Text119.RowSource = "" Then
        Me.Text119.RowSource = "SELECT [category] FROM " & ConSQL & ".[CCM." & MST15 & "] ORDER BY [category];"
        End If
End Sub
        
Private Sub Text120_GotFocus()
        If Me.Text120.RowSource = "" Then
        Me.Text120.RowSource = "SELECT [Initiative_ID] FROM " & ConSQL & ".[CCM." & MST08 & "] WHERE [use] = true;"
        End If
End Sub

Private Sub Text122_GotFocus()
        If Me.Text122.RowSource = "" Then
        Me.Text122.RowSource = "SELECT DISTINCT [Department] FROM " & ConSQL & ".[CCM." & MST05 & "] ORDER BY [Department];"
        End If
End Sub

Private Sub Text123_GotFocus()
        If Me.Text123.RowSource = "" Then
        Me.Text123.RowSource = "SELECT DISTINCT [service_category] FROM " & ConSQL & ".[CCM." & MST05 & "] ORDER BY [service_category];"
        End If
End Sub

Private Sub Text129_GotFocus()
        If Me.Text129.RowSource = "" Then
        Me.Text129.RowSource = "SELECT DISTINCT [code] FROM " & ConSQL & ".[CCM." & MST09 & "] ORDER BY [code];"
        End If
End Sub

Private Sub Text141_GotFocus()
        If Me.Text141.RowSource = "" Then
        Me.Text141.RowSource = "SELECT DISTINCT [year] FROM " & ConSQL & ".[CCM." & MST16 & "] ORDER BY [year];"
        End If
End Sub

Private Sub Text206_GotFocus()
        If Me.Text206.RowSource = "" Then
        Me.Text206.RowSource = "SELECT [name] FROM " & ConSQL & ".[CCM." & MST04 & "] WHERE [use] = 1 ORDER BY [ID];"
        End If
End Sub

Private Sub Text207_GotFocus()
        If Me.Text207.RowSource = "" Then
        Me.Text207.RowSource = "SELECT DISTINCT [name] FROM " & ConSQL & ".[CCM." & MST12 & "];"
        End If
End Sub

Private Function pf_Clear_Combobox()
        On Error Resume Next
        Dim meControl As Control
        For Each meControl In Me.Controls
                If meControl.ControlType = acComboBox Then
                        meControl.RowSource = ""
                End If
        Next
        On Error GoTo 0
End Function

Private Function pf_Reset_Contract_Info()
        Call f_Choose_Item(contractID)
        Call f_Choose_Extra(contractID)
        Call f_List_Price(contractID)
        Call f_List_Attachment(contractID)
End Function

Private Function pf_Button_Reset()
        Call pf_Button_Enable_Contract(1, 0, 1, 1)
        Call pf_Button_Enable_Price(0, 0, 1, 0)
        Call pf_Button_Enable_Attachment(1, 1, 1)
        Call pf_Button_Initialize_Contract
        Call pf_Button_Initialize_Price
End Function

Private Function pf_Button_Enable_Contract(ed As Boolean, up As Boolean, ne As Boolean, co As Boolean)
        Me.ComContractEdit.Enabled = ed
        Me.ComContractUpdate.Enabled = up
        Me.ComContractNew.Enabled = ne
        Me.ComContractCopy.Enabled = co
End Function

Public Function pf_Button_Enable_Price(ed As Boolean, up As Boolean, ne As Boolean, co As Boolean)
        Me.ComPriceEdit.Enabled = ed
        Me.ComPriceUpdate.Enabled = up
        Me.ComPriceNew.Enabled = ne
        Me.ComPriceCopy.Enabled = co
End Function

Private Function pf_Button_Enable_Attachment(ad As Boolean, dt As Boolean, dl As Boolean)
        Me.ComAttachAdd.Enabled = ad
        Me.ComAttachDelete.Enabled = dt
        Me.ComAttachDownload.Enabled = dl
End Function

Private Function pf_Button_Initialize_Contract()
        Me.ComContractEdit.Caption = "Edit"
        Me.ComContractUpdate.Caption = "Update"
        Me.ComContractNew.Caption = "New"
        Me.ComContractCopy.Caption = "Copy to New"
End Function

Private Function pf_Button_Initialize_Price()
        Me.ComPriceEdit.Caption = "Edit"
        Me.ComPriceUpdate.Caption = "Update"
        Me.ComPriceNew.Caption = "New"
        Me.ComPriceCopy.Caption = "Copy to New"
End Function

Private Function pf_Authentication_FRM03()
        Dim comButton As Control
        Dim buttonLevel As Integer
        With Me.Form
                For Each comButton In .Controls
                        If comButton.ControlType = acCommandButton Then
                                Debug.Print comButton.Name & ": " & authLevel & " - " & comButton.Properties("tag")
                                buttonLevel = comButton.Properties("tag")
                                If authLevel - buttonLevel >= 0 Then
                                        comButton.Enabled = True
                                        comButton.Visible = True
                                End If
                                If authLevel - buttonLevel < 0 Then
                                        comButton.Enabled = False
                                        comButton.Visible = True
                                End If
                                If authLevel - buttonLevel < -5 Then
                                        comButton.Enabled = False
                                        comButton.Visible = False
                                End If
                        End If
                Next
        End With
                
        Select Case authProperty
                Case "M"
                        Call pf_Detail_Control_Enabled(True):   Call pf_Detail_Control_Visible(True)
                Case "W"
                        Call pf_Detail_Control_Enabled(True):   Call pf_Detail_Control_Visible(True)
                Case "R"
                        Call pf_Detail_Control_Enabled(False):  Call pf_Detail_Control_Visible(True)
                Case Else
                        Call pf_Detail_Control_Enabled(False):  Call pf_Detail_Control_Visible(False)
        End Select
        Select Case authCost
                Case "M"
                        Call pf_Price_Control_Enabled(True):    Call pf_Price_Control_Visible(True)
                Case "W"
                        Call pf_Price_Control_Enabled(True):    Call pf_Price_Control_Visible(True)
                Case "R"
                        Call pf_Price_Control_Enabled(False):   Call pf_Price_Control_Visible(True)
                Case Else
                        Call pf_Price_Control_Enabled(False):   Call pf_Price_Control_Visible(False)
        End Select
        Select Case authOther
                Case "M"
                        Call pf_Other_Control_Enabled(True):    Call pf_Other_Control_Visible(True)
                Case "W"
                        Call pf_Other_Control_Enabled(True):    Call pf_Other_Control_Visible(True)
                Case "R"
                        Call pf_Other_Control_Enabled(False):   Call pf_Other_Control_Visible(True)
                Case Else
                        Call pf_Other_Control_Enabled(False):   Call pf_Other_Control_Visible(False)
        End Select
EX:
End Function

Private Function pf_Detail_Control_Visible(vis As Boolean)
        On Error Resume Next
        Dim i As Integer
        For i = 2 To 33
                Me.Controls("Label" & 100 + i).Visible = vis:   Me.Controls("Text" & 100 + i).Visible = vis
        Next
        On Error GoTo 0
End Function

Private Function pf_Price_Control_Visible(vis As Boolean)
        On Error Resume Next
        Dim i As Integer
        For i = 3 To 11
                Me.Controls("Label" & 200 + i).Visible = vis:   Me.Controls("Text" & 200 + i).Visible = vis
        Next
        On Error GoTo 0
End Function

Private Function pf_Other_Control_Visible(vis As Boolean)
        Dim i As Integer
        On Error Resume Next
        For i = 1 To 1
                Me.Controls("Label" & 300 + i).Visible = vis:   Me.Controls("Text" & 300 + i).Visible = vis
        Next
        On Error GoTo 0
End Function

Private Function pf_Detail_Control_Enabled(vis As Boolean)
        Dim i As Integer
        On Error Resume Next
        For i = 4 To 32
                Me.Controls("Text" & 100 + i).Enabled = vis
        Next
        For i = 1 To 5
                Me.Controls("Text" & 150 + i).Enabled = vis
        Next
        On Error GoTo 0
End Function

Private Function pf_Price_Control_Enabled(vis As Boolean)
        Dim i As Integer
        On Error Resume Next
        For i = 3 To 11
                Me.Controls("Text" & 200 + i).Enabled = vis
        Next
        On Error GoTo 0
End Function

Private Function pf_Other_Control_Enabled(vis As Boolean)
        Dim i As Integer
        On Error Resume Next
        For i = 1 To 1
                Me.Controls("Text" & 300 + i).Enabled = vis
        Next
        On Error GoTo 0
End Function

Private Function pf_Button_Cancel(comButton As Control) As Boolean
        If comButton.Caption <> "Cancel" Then
                pf_Button_Cancel = False
                comButton.Caption = "Cancel"
        Else
                pf_Button_Cancel = True
                comButton.Caption = comButton.Properties("controltiptext")
        End If
End Function

Private Function pf_ID_Store(i As Integer)
        Select Case i
                Case 0
                        contractID = tmpContractID
                        contractNumber = tmpContractNumber
                        contractStatus = tmpContractStatus
                        tmpContractID = 0
                        tmpContractNumber = ""
                        tmpContractStatus = ""
                Case 1
                        tmpContractID = contractID
                        tmpContractNumber = contractNumber
                        tmpContractStatus = contractStatus
                        contractID = 0
                        contractNumber = ""
                        contractStatus = ""
        End Select
End Function
