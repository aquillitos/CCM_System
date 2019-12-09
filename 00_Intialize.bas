Attribute VB_Name = "00_Intialize"
Option Compare Database:    Option Explicit

Public db As DAO.Database
'===== Database Connect ==========
Public ConSys As ADODB.Connection 'Required "Microsoft ActiveX Data Objects 6.1(or more) Library"
Public ConSQL As String    'ADODB SQL Connection String
Public ConADO As String    'ADODB ADO Connection String
Public ConLo As ADODB.Connection

'===== Froms ===================
Public FRM01 As String
Public FRM02 As String
Public FRM03 As String
Public FRM04 As String
Public FRM05 As String
Public FRM06 As String

'===== CCM Data Table ===========
Public CCMDATA As String
Public CCMCOST1 As String
Public CCMCOST2 As String
Public CCMBOID As String
Public CCMAttach As String
Public CCMHistory As String

'===== Master Data Table ==========
Public MST00 As String
Public MST01 As String, MST02 As String, MST03 As String, MST04 As String, MST05 As String
Public MST06 As String, MST07 As String, MST08 As String, MST10 As String
Public MST09 As String, MST12 As String, MST13 As String, MST14 As String, MST15 As String
Public MST16 As String

'===== Temporary Table ===========
Public TMP01 As String
Public TMP11 As String
Public TMP21 As String, TMP22 As String, TMP23 As String, TMP24 As String
Public TMP31 As String, TMP32 As String, TMP33 As String
Public TMP51 As String, TMP52 As String, TMP53 As String, TMP54 As String

'===== User Authentication =========
Public userID As String
Public userName As String
Public userMail As String
Public authName As String
Public authMaster As String
Public authentication_level As Integer
Public authProperty As String
Public authCost As String
Public authOther As String
Public authCase1 As String

'===== Parameter Data ============
Public contractID As Long
Public contractNumber As String
Public contractStatus As String
Public priceID As Long
Public tmpContractID As Long
Public tmpContractNumber As String
Public tmpContractStatus As String
Public authLevel As Integer

Public savID As Double
Public savCID As Double
Public savNum As String
Public savVer As Long
Public savSta As Date
Public savEnd As Date
Public savInS As Boolean
Public savInE As Boolean
Public savMon As Integer
Public savCur As String
Public savTer As String
Public savAmo As Currency
Public savMoC As Currency
Public savAnC As Currency
        
Public ANS As String

Public fieldCount As Integer
Public pUnit As String

Public cNumPref1 As String, cNumPref2 As String
Public cNumPref3 As String, cNumPref4 As String
Public cNumPref5 As String, cNumPref6 As String

'===== File Management ============
Public DeskTopPath As String
Public filePath As String
Public attachDir As String
Public fileindex As Integer

'===== Application Form Management ==
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public RectFrm As RECT
Public RectAcc As RECT
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Function f_Initialize()

        '=== From Name =========================================
        FRM01 = "01_Main"
        FRM03 = "03_Detail"
        FRM04 = "04_Detail_Finance"
        FRM05 = "04_Master_Maintenance"
        FRM06 = "06_Data_Maintenance"
        
        '=== Master Table Name ===================================
        MST00 = "CCM_MST_System_Config"
        MST01 = "CCM_MST_User"
        MST02 = "CCM_MST_Authentication"
        MST03 = "CCM_MST_Status"
        MST04 = "CCM_MST_Term"
        MST05 = "CCM_MST_Service"
        MST06 = "CCM_MST_Contract_Type"
        MST07 = "CCM_MST_Vendor"
        MST08 = "CCM_MST_Initiative_ID"
        MST09 = "CCM_MST_Allocation_Code"
        MST12 = "CCM_MST_Currency"
        MST13 = "CCM_MST_Address"
        MST14 = "CCM_MST_Budget_Number"
        MST15 = "CCM_MST_Investment_Category"
        MST16 = "CCM_MST_BOID"
        
        '=== Temporary Table =====================================
        TMP01 = "tmp_Payment_Schedule"
        TMP11 = "tmp_allocation_code"
        TMP21 = "tmp_Download_Record"
        TMP22 = "tmp_Download_Price"
        TMP23 = "tmp_Download_BOID"
        TMP24 = "tmp_Download"
        
        TMP31 = "tmp_Import_Error"
        TMP32 = "tmp_Import_CCM"
        TMP33 = "tmp_Import_Price"
        
        TMP51 = "tmp_Contract"
        TMP52 = "tmp_Cost"
        TMP53 = "tmp_Cost2"
        TMP54 = "tmp_BOID"
        
        '=== Record Data Table  ===================================
        CCMDATA = "CCM_Data"
        CCMCOST1 = "CCM_Data_Cost_Term"
        CCMCOST2 = "CCM_Data_Cost_Detail"
        CCMAttach = "CCM_Data_Attachment"
        CCMHistory = "CCM_Data_Update_History"
        CCMBOID = "CCM_Data_BOID"
        
        '=== Prefix for Contract Number Naming ========================
        cNumPref1 = "CNTRC"
        cNumPref2 = "CNTRL"
        cNumPref3 = "CNTRM"
        cNumPref4 = "SRV"
        cNumPref5 = "SRV"
        cNumPref6 = "ZZZ"
        
        '=== Number of Textbox, Combobox in Main Form(FRM01) ===========
        fieldCount = 16
        
        
        Set db = CurrentDb
        Call s_Get_System_Config
        Call s_System_Server_Connect
        Call s_Login_ID
        Call s_DeskTopPath
        Call s_System_Config_Update
        Call s_Local_Server_Connect
End Function
