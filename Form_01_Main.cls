VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_01_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database:    Option Explicit

Private Sub Form_Load()
        'DoCmd.ShowToolbar "Ribbon", acToolbarNo
        'DoCmd.SelectObject acForm, FRM01, True
        'DoCmd.RunCommand acCmdWindowHide
        
        Call f_Initialize
        Call f_Cancelling
        Call pf_Clear_Combobox
        Call pf_List_Zero
        Call pf_Authentication_FRM01
        Call pf_Count_SearchIndex
End Sub

Private Sub Form_Unload(Cancel As Integer)
        Call pf_Clear_Combobox
        Call f_Delete_MasterQuery
        Call f_Delete_Temporary
End Sub

Private Sub Form_Close()
        Call pf_Clear_Combobox
        Call f_Delete_MasterQuery
        Call f_Delete_Temporary
End Sub

Private Sub ComInitial_Click()
        Call f_Initialize
End Sub

'===== Command Buttons =================================================================================
Private Sub ComSearch_Click()
        DoCmd.Hourglass True
        Call f_List_Contract(fieldCount)
        DoCmd.Hourglass False
End Sub

Private Sub ComReset_Click()
        Dim i As Integer
        Dim controlName As String
        For i = 1 To fieldCount
                controlName = "Search" & Right("00" & i, 3)
                Me.Controls(controlName) = ""
        Next
        Call pf_List_Zero
End Sub

Private Sub ComCreateNew_Click()
        contractID = 0
        contractNumber = ""
        DoCmd.Close acForm, FRM03
        DoCmd.OpenForm FRM03, acNormal
End Sub

Private Sub ComExit_Click()
        Call f_Delete_MasterQuery
        Call pf_Clear_Combobox
        Application.Quit
End Sub

Private Sub ComMaintenance_Click()
        DoCmd.Minimize
        DoCmd.OpenForm FRM05, acNormal
End Sub

Private Sub ComMaster_Click()
        DoCmd.Close acForm, FRM01
        DoCmd.OpenForm FRM05, acNormal
End Sub

Private Sub ComDwonload_Click()
        If f_Download(fieldCount) = True Then
                ANS = MsgBox("Exported onto your desktop.", vbInformation + vbOKOnly)
        End If
End Sub

'===== ComboBox List ============================================================================
Private Sub Search001_GotFocus()
        If Me.Search001.RowSource = "" Then
        Me.Search001.RowSource = "SELECT DISTINCT [status] FROM " & ConSQL & ".[CCM." & MST03 & "] WHERE [use] = true;"
        End If
End Sub

Private Sub Search002_GotFocus()
        If Me.Search002.RowSource = "" Then
        Me.Search002.RowSource = "SELECT DISTINCT [Initiative_ID] FROM " & ConSQL & ".[CCM." & MST08 & "] ORDER BY [Initiative_ID];"
        End If
End Sub

Private Sub Search003_GotFocus()
        If Me.Search003.RowSource = "" Then
        Me.Search003.RowSource = "SELECT DISTINCT [Initiative_Name] FROM " & ConSQL & ".[CCM." & MST08 & "] ORDER BY [Initiative_Name];"
        End If
End Sub

Private Sub Search004_GotFocus()
        If Me.Search004.RowSource = "" Then
        Me.Search004.RowSource = "SELECT DISTINCT [Department] FROM " & ConSQL & ".[CCM." & MST05 & "] ORDER BY [Department];"
        End If
End Sub

Private Sub Search005_GotFocus()
        If Me.Search005.RowSource = "" Then
        Me.Search005.RowSource = "SELECT DISTINCT [service_category] FROM " & ConSQL & ".[CCM." & MST05 & "] ORDER BY [service_category];"
        End If
End Sub

Private Sub Search006_GotFocus()
        If Me.Search006.RowSource = "" Then
        Me.Search006.RowSource = "SELECT DISTINCT [contract_model] FROM " & ConSQL & ".[CCM." & MST06 & "];"
        End If
End Sub

Private Sub Search007_GotFocus()
        If Me.Search007.RowSource = "" Then
        Me.Search007.RowSource = "SELECT DISTINCT [contract_type] FROM " & ConSQL & ".[CCM." & MST06 & "];"
        End If
End Sub

Private Sub Search008_GotFocus()
        If Me.Search008.RowSource = "" Then
        Me.Search008.RowSource = "SELECT DISTINCT [vendor] FROM " & ConSQL & ".[CCM." & MST07 & "] WHERE [use] = true;"
        End If
End Sub
        
Private Sub Search010_GotFocus()
        If Me.Search010.RowSource = "" Then
        Me.Search010.RowSource = "SELECT [eDate] FROM (SELECT DISTINCT FORMAT([end],'yyyy-mm') AS [eDate] FROM " & ConSQL & ".[CCM." & CCMDATA & "]) ORDER BY [eDate] DESC;"
        End If
End Sub

Private Sub Search015_GotFocus()
        If Me.Search015.RowSource = "" Then
        Me.Search015.RowSource = "SELECT [sDate] FROM (SELECT DISTINCT FORMAT([start],'yyyy-mm') AS [sDate] FROM " & ConSQL & ".[CCM." & CCMDATA & "]) ORDER BY [sDate] DESC;"
        End If
End Sub

Private Sub Search016_GotFocus()
        If Me.Search016.RowSource = "" Then
        Me.Search016.RowSource = "SELECT [cDate] FROM (SELECT DISTINCT FORMAT([cancel_date],'yyyy-mm') AS [cDate] FROM " & ConSQL & ".[CCM." & CCMDATA & "]) ORDER BY [cDate] DESC;"
        End If
End Sub


'===== Private Functions ===================================================================
Private Function pf_Count_SearchIndex()
        On Error Resume Next
        Dim meControl As Control, c As Integer
        c = 0
        For Each meControl In Me.Controls
                With meControl
                        If .Name Like "Search*" Then
                                c = c + 1
                        End If
                End With
        Next
        fieldCount = c: c = 0
End Function

Private Function pf_Authentication_FRM01()
        On Error Resume Next
        Dim meControl As Control
        Dim eStatus As Boolean, vStatus As Boolean
                        
        For Each meControl In Me.Controls
                With meControl
                        If .ControlType = acCommandButton Then
                                Debug.Print .Name & " : " & .Properties("HelpContextId")
                                Select Case authMaster
                                        Case "M"
                                                If .Properties("HelpContextId") >= 0 Then
                                                        .Enabled = True:    .Visible = True
                                                End If
                                        Case "W"
                                                If .Properties("HelpContextId") >= 5 Then
                                                        .Enabled = True:    .Visible = True
                                                Else
                                                        .Enabled = False:    .Visible = False
                                                End If
                                        Case Else
                                                If .Properties("HelpContextId") >= 10 Then
                                                        .Enabled = True:    .Visible = True
                                                Else
                                                        .Enabled = False:    .Visible = False
                                                End If
                                End Select
                        End If
                End With
        Next
End Function

Private Sub pf_List_Zero()
        Dim stSQL1 As String
        stSQL1 = "SELECT * FROM " & ConSQL & ".[CCM." & CCMDATA & "] WHERE [ID] = 0"
        Forms(FRM01).SUB001.Form.RecordSource = stSQL1: stSQL1 = ""
        Forms(FRM01).Text101 = 0
End Sub

Private Function pf_Clear_Combobox()
        Dim i As Integer
        On Error Resume Next
        For i = 1 To fieldCount
                Me.Controls("Search" & Right("000" & i, 3)).RowSource = ""
        Next i
        i = 0
        On Error GoTo 0
End Function
