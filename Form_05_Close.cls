VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_05_Close"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ComCloseTable_Click()
        DoCmd.Close acTable, tName, acSaveYes
        DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
