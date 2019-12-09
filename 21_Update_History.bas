Attribute VB_Name = "21_Update_History"
Option Compare Database:    Option Explicit

Public Function f_Update_History(orgValue, newValue, targetTable, targetField)
        '�ۑ�����Ă���f�[�^�ƈقȂ�f�[�^�������ݎg�p�Ƃ����ꍇ�A�����Ƃ��Ďc���B
        If orgValue = newValue Then '�f�[�^��������������
                'Do nothing
        Else    '�f�[�^���قȂ��Ă�����
                If Nz(newValue, "") <> "" And targetField <> "updated_on" Then
                        Dim stSQL9 As String
                        Dim rs9 As ADODB.Recordset
                        Set rs9 = New ADODB.Recordset
                        
                        stSQL9 = "SELECT * FROM CCM." & CCMHistory & ";"
                        rs9.Open stSQL9, ConSys, adOpenDynamic, adLockOptimistic
                        rs9.AddNew
                                rs9![contract_ID] = contractID  'ID
                                rs9![contract_Number] = contractNumber  'Contract Number
                                rs9![Table] = targetTable   'Table
                                rs9![Field] = targetField   'Field
                                rs9![old] = orgValue    'Old Value
                                rs9![New] = newValue    'New Value
                                rs9![updated_on] = Now()    'Updated Data-Time
                                rs9![updated_by] = userName 'Updated Person
                        rs9.Update
                        rs9.Close:  Set rs9 = Nothing:  stSQL9 = ""
                End If
        End If
End Function
