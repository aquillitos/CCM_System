Attribute VB_Name = "52_Status_Change"
Option Compare Database:    Option Explicit

Public Function f_Cancelling()
        Dim stSQL1 As String
        Dim rs1 As ADODB.Recordset
        Dim contractIDOriginal As String
        Dim userNameOriginal As String
        Dim cancelStatus As String
        
        Set rs1 = New ADODB.Recordset
            
        cancelStatus = "Active"
        stSQL1 = "SELECT [ID], [status], [cancel_date] FROM CCM.[" & CCMDATA & "] WHERE [status] = '" & cancelStatus & "' AND [cancel_date] <= " & DateSerial(Year(Now()), Month(Now()), Day(Now())) & " ;"
        rs1.Open stSQL1, ConSys, adOpenDynamic, adLockOptimistic
                Do Until rs1.EOF
                        rs1![status] = "Cancelled"
                        rs1.Update
                        userNameOriginal = userName
                        userName = "System"
                        contractIDOriginal = contractID
                        contractID = rs1![ID]
                        Call f_Update_History("Active", "Cancelled", CCMDATA, "status")
                        userName = userNameOriginal
                        contractID = contractIDOriginal
                        rs1.MoveNext
                Loop
        rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
        userNameOriginal = ""
        contractIDOriginal = ""
End Function

