Attribute VB_Name = "15_Extra_Item"
Option Compare Database:    Option Explicit

Public Function f_Choose_Extra(cID)
        Dim stSQL1 As String:   Dim rs1 As ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        stSQL1 = "SELECT [ID], [contract_model], [contract_Type] FROM CCM.[" & CCMDATA & "] WHERE [ID] = " & cID & ";"
        rs1.Open stSQL1, ConSys, adOpenDynamic, adLockReadOnly
                If Not rs1.EOF Then
                        Call f_Show_Extra(Nz(rs1![contract_model], ""), Nz(rs1![contract_type], ""))
                End If
        rs1.Close:  Set rs1 = Nothing:   stSQL1 = ""
End Function

Public Function f_Show_Extra(cModel As String, cType As String)
        Dim stSQL1 As String:   Dim rs1 As ADODB.Recordset
        Dim i As Integer
        Set rs1 = New ADODB.Recordset
        
        stSQL1 = "SELECT * FROM CCM.[" & MST06 & "] WHERE [contract_model] = '" & cModel & "' AND [contract_type] = '" & cType & "';"
        rs1.Open stSQL1, ConSys, adOpenDynamic, adLockReadOnly
                With Forms(FRM03)
                If Not rs1.EOF Then
                            If Nz(rs1![Extra_field_1], "") <> "" Then
                                    .Label151.Visible = True:   .Label151.Caption = rs1![Extra_field_1]
                                    .Text151.Visible = True:    .Text151 = Format(.Text151, pf_Set_Attribute(rs1![Extra_field_1_Attribute]))
                            Else
                                    .Label151.Visible = False:   .Text151.Visible = False
                                    .Label151.Caption = "*Extra_Field_1"
                            End If
                            If Nz(rs1![Extra_field_2], "") <> "" Then
                                    .Label152.Visible = True:   .Label152.Caption = rs1![Extra_field_2]
                                    .Text152.Visible = True:    .Text152 = Format(.Text152, pf_Set_Attribute(rs1![Extra_field_2_Attribute]))
                            Else
                                    .Label152.Visible = False:   .Text152.Visible = False
                                    .Label152.Caption = "*Extra_Field_2"
                            End If
                            If Nz(rs1![Extra_field_3], "") <> "" Then
                                    .Label153.Visible = True:   .Label153.Caption = rs1![Extra_field_3]
                                    .Text153.Visible = True:    .Text153 = Format(.Text153, pf_Set_Attribute(rs1![Extra_field_3_Attribute]))
                            Else
                                    .Label153.Visible = False:   .Text153.Visible = False
                                    .Label153.Caption = "*Extra_Field_3"
                            End If
                            If Nz(rs1![Extra_field_4], "") <> "" Then
                                    .Label154.Visible = True:   .Label154.Caption = rs1![Extra_field_4]
                                    .Text154.Visible = True:    .Text154 = Format(.Text154, pf_Set_Attribute(rs1![Extra_field_4_Attribute]))
                            Else
                                    .Label154.Visible = False:   .Text154.Visible = False
                                    .Label154.Caption = "*Extra_Field_4"
                            End If
                            If Nz(rs1![Extra_field_5], "") <> "" Then
                                    .Label155.Visible = True:   .Label155.Caption = rs1![Extra_field_5]
                                    .Text155.Visible = True:    .Text155 = Format(.Text155, pf_Set_Attribute(rs1![Extra_field_5_Attribute]))
                            Else
                                    .Label155.Visible = False:   .Text155.Visible = False
                                    .Label155.Caption = "*Extra_Field_5"
                            End If
                Else
                        For i = 1 To 5
                                .Controls("Label" & 150 + i).Visible = False:   .Controls("Text" & 150 + i).Visible = False
                                .Controls("Label" & 150 + i).Caption = "*extra_Field_" & i
                                .Controls("Text" & 150 + i) = ""
                        Next i
                End If
                End With
        rs1.Close:  Set rs1 = Nothing:  stSQL1 = ""
End Function

Private Function pf_Set_Attribute(fAttribute) As String
        Select Case fAttribute
                Case "Text"
                        pf_Set_Attribute = "@"
                Case "Date"
                        pf_Set_Attribute = "yyyy-mm-dd"
                Case "Double"
                        pf_Set_Attribute = "#,##0.00"
                Case "Currency"
                        pf_Set_Attribute = "#,##0"
                Case "Number"
                        pf_Set_Attribute = "#,##0"
                Case "Long"
                        pf_Set_Attribute = "#,##0"
                Case "Single"
                        pf_Set_Attribute = "#,##0"
                Case Else
                        pf_Set_Attribute = fAttribute
        End Select
End Function
