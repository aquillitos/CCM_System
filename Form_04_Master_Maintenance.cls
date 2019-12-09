VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_04_Master_Maintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database:    Option Explicit

'===== Form Control ===========================================================================================================
'
'==========================================================================================================================
Private Sub ComClose_Click()
        DoCmd.Close acForm, FRM05
        DoCmd.OpenForm FRM01, acNormal
End Sub

Private Sub Form_Load()
        Call f_Initialize_Admin
        Dim rs1 As ADODB.Recordset
        Dim nDate As Double
        
        Dim thisControl As Control
        For Each thisControl In Me.Controls
                Select Case thisControl.ControlType
                        Case acComboBox
                                thisControl.RowSource = ""
                        Case acTextBox
                                thisControl = ""
                End Select
        Next
        
        Me.Combo101.RowSource = "SELECT [TABLE_NAME] FROM " & ConSQL & ".[INFORMATION_SCHEMA.TABLES] WHERE [TABLE_NAME] LIKE 'CCM_MST_*' ORDER BY [TABLE_NAME];"
        Me.Combo102.RowSource = "SELECT [TABLE_NAME] FROM " & ConSQL & ".[INFORMATION_SCHEMA.TABLES] WHERE [TABLE_NAME] LIKE 'CCM_Data*' OR [TABLE_NAME] LIKE 'RPT_*' ORDER BY [TABLE_NAME];"
        Me.Combo301.RowSource = "SELECT [TABLE_NAME] FROM " & ConSQL & ".[INFORMATION_SCHEMA.TABLES] ORDER BY [TABLE_NAME];"
        Me.Combo401.RowSource = "SELECT [TABLE_NAME] FROM " & ConSQL & ".[INFORMATION_SCHEMA.TABLES] WHERE [TABLE_NAME] LIKE 'CCM_Data*' OR [TABLE_NAME] LIKE 'RPT_*' ORDER BY [TABLE_NAME];"
        
        Me.Combo321.RowSource = "<; <=; =; >=; >; <>; IS; LIKE"
        Me.Combo322.RowSource = "<; <=; =; >=; >; <>; IS; LIKE"
        Me.Combo323.RowSource = "<; <=; =; >=; >; <>; IS; LIKE"
        
        
        nDate = Now()
        Dim m As Integer
        For m = 36 To -36 Step -1
                Me.Combo201.AddItem DateAdd("m", m, nDate) & ";" & Format(DateAdd("m", m, nDate), "YYYY-MM")
        Next m
        Me.Combo201 = DateAdd("m", 0, nDate)
End Sub

Private Sub Form_Unload(Cancel As Integer)
        Dim mTable As TableDef
        For Each mTable In db.TableDefs
            If mTable.Name <> MST00 Then
                    'Call f_Delete_LocalTable(mTable.Name)
            End If
        Next
End Sub


'===== Budget Report Creation ===================================================================================================
'
'==========================================================================================================================
Private Sub Combo201_Change()
        Dim sDate As Date, ssDate As Double
        sDate = Me.Combo201
        ssDate = sDate
        Me.Combo202.RowSource = ""
        Dim m As Integer
        For m = 48 To 0 Step -1
                Me.Combo202.AddItem DateAdd("m", m, ssDate) & ";" & Format(DateAdd("m", m, ssDate), "YYYY-MM")
        Next m
        Me.Combo202 = sDate
End Sub

Private Sub ComReport_Click()
        ANS = MsgBox("Do you want to proceed?" & vbCrLf & "This will take some time to finsh.", vbQuestion + vbYesNo)
        If ANS = vbYes Then
                DoCmd.Hourglass True
                Dim sDate As Double, eDate As Double
                Debug.Print Me.Combo201 & " : " & Me.Combo202
                sDate = DateSerial(Year(Me.Combo201), Month(Me.Combo201), 1)
                eDate = DateSerial(Year(Me.Combo202), Month(Me.Combo202), 1)
                Call f_Export_Report(sDate, eDate)
                DoCmd.Hourglass False
        End If
End Sub


'===== Master Data Maintenance =================================================================================================
'
'==========================================================================================================================
Private Sub Combo101_Change()
        DoCmd.Close acTable, tName1, acSaveYes
        tName1 = Me.Combo101
End Sub

Private Sub ComTableOpen_Click()
        If Nz(tName1, "") <> "" Then
                Call f_Delete_LocalTable(tName1)
                Call f_Delete_LocalTable(tName2)
                Call f_Copy_TableSL(tName1, tName1)
                If f_Exist_LocalTable(tName1) = True Then
                        DoCmd.OpenTable tName1
                End If
        End If
End Sub

Private Sub ComTableUpload_Click()
        If tName1 <> "dummy" Then
                ANS = MsgBox("Do you want to update the Master Data?" & vbCrLf & vbCrLf & _
                                    "This process ..." & vbCrLf & _
                                    "!!... DELETE the Master Data from the Server and ..." & vbCrLf & _
                                    "!!... UPLOAD the Current Data into the Server." & vbCrLf & vbCrLf & _
                                    "YOU CANNOT ROLLBACK to the previous state once completed!!!", vbExclamation + vbYesNo)
                If ANS = vbYes Then
                        'HERE UPLOAD PROCEDURE
                        Call f_Upsert(tName1)
                        
                        ANS = MsgBox("Uploaded", vbInformation + vbOKOnly)
                Else
                        ANS = MsgBox("Upload Cancelled.", vbInformation + vbOKOnly)
                End If
        End If
End Sub

Private Sub ComCloseTable_Click()
        DoCmd.Close acTable, tName1, acSaveYes
        Me.Combo101 = ""
        tName1 = "dummy"
End Sub


'===== Record Data Maintenance =================================================================================================
'
'==========================================================================================================================
Private Sub Combo102_Change()
        DoCmd.Close acTable, tName2, acSaveYes
        tName2 = Me.Combo102
End Sub

Private Sub ComTableOpen2_Click()
        If Nz(tName2, "") <> "" Then
                Call f_Delete_LocalTable(tName1)
                Call f_Delete_LocalTable(tName2)
                Call f_Copy_TableSL(tName2, tName2)
                If f_Exist_LocalTable(tName2) = True Then
                        DoCmd.OpenTable tName2
                End If
        End If
End Sub

Private Sub ComTableUpload2_Click()
        If tName2 <> "dummy" Then
                ANS = MsgBox("Do you want to update the Master Data?" & vbCrLf & vbCrLf & _
                                    "This process ..." & vbCrLf & _
                                    "!!... DELETE the Master Data from the Server and ..." & vbCrLf & _
                                    "!!... UPLOAD the Current Data into the Server." & vbCrLf & vbCrLf & _
                                    "YOU CANNOT ROLLBACK to the previous state once completed!!!", vbExclamation + vbYesNo)
                If ANS = vbYes Then
                        'HERE UPLOAD PROCEDURE
                        ANS = MsgBox("Uploaded", vbInformation + vbOKOnly)
                Else
                        ANS = MsgBox("Upload Cancelled.", vbInformation + vbOKOnly)
                End If
        End If
End Sub

Private Sub ComCloseTable2_Click()
        DoCmd.Close acTable, tName2, acSaveYes
        Me.Combo102 = ""
        tName2 = "dummy"
End Sub

'===== Data Maintenance By SQL ================================================================================================
'
'==========================================================================================================================
Private Sub Combo301_Change()
        Dim stSQL1 As String
        stSQL1 = "SELECT [COLUMN_NAME], [DATA_TYPE] FROM " & ConSQL & ".[INFORMATION_SCHEMA.COLUMNS] WHERE [TABLE_NAME] = '" & Me.Combo301 & "' AND [COLUMN_NAME] <> 'ID' ;"
        Me.Combo311.RowSource = stSQL1
        Me.Combo312.RowSource = stSQL1
        Me.Combo313.RowSource = stSQL1
        Me.Combo351.RowSource = stSQL1
        stSQL1 = ""
End Sub

Private Sub ComSetSQL_Click()
        Call f_Create_SQL_String
End Sub

Private Sub ComRunSQL_Click()
        ANS = MsgBox(f_Check_NumberOfRecord(Me.Text362) & " Record will be updated." & vbCrLf & vbCrLf & _
                                "Do you want to RUN UPDATE query?" & vbCrLf & _
                                "This will over-write the data in DATABASE Server." & vbCrLf & _
                                "And CANNOT ROLLBACK.", vbExclamation + vbYesNo)
        If ANS = vbYes Then
                Call f_Run_UpdateQuery(Me.Text361)
        End If
End Sub

'===== Data Bulk Upload =======================================================================================================
'
'==========================================================================================================================
Private Sub ComExport_Click()
                
End Sub

Private Sub ComUpload_Click()
        Call f_Bulk_Upload(Me.Combo401)
End Sub

Private Sub Text311_AfterUpdate()
        If Me.Combo321 = "LIKE" Then
                Me.Text311 = "%" & Me.Text311 & "%"
        End If
End Sub

Private Sub Text312_AfterUpdate()
        If Me.Combo322 = "LIKE" Then
                Me.Text312 = "%" & Me.Text312 & "%"
        End If
End Sub

Private Sub Text313_AfterUpdate()
        If Me.Combo323 = "LIKE" Then
                Me.Text313 = "%" & Me.Text313 & "%"
        End If
End Sub

Private Sub Text361_AfterUpdate()
        Dim oSQL As String
        Dim wSQL As String
        Dim uSQL As String
        oSQL = Me.Text361
        
        uSQL = Replace(Left(oSQL, InStr(1, oSQL, "SET") - 1), "UPDATE", "SELECT COUNT([ID]) AS [IDS] FROM")
        wSQL = Right(oSQL, Len(oSQL) - InStr(1, oSQL, "WHERE") + 1)
        
        Me.Text362 = uSQL & wSQL
End Sub
