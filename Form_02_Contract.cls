VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_02_Contract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database:    Option Explicit

Private Sub Form_DblClick(Cancel As Integer)
        DoCmd.Hourglass True
        On Error GoTo ERR
        If IsNumeric(Me.ID) = True Then
                contractID = Me.ID
                contractNumber = Me.number
                contractStatus = Me.status
                
                'Select Case authName
                '        Case "User2"        'Load UI for Finance (Currently Disabled)
                                'DoCmd.Close acForm, FRM04
                                'DoCmd.OpenForm FRM04
                '        Case Else
                                DoCmd.Close acForm, FRM03
                                DoCmd.OpenForm FRM03
                'End Select
        End If
ERR:
        DoCmd.Hourglass False
End Sub
