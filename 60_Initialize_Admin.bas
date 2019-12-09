Attribute VB_Name = "60_Initialize_Admin"
Option Compare Database:    Option Explicit

Public tName1 As String, tName2 As String

Public Function f_Initialize_Admin()
        Call f_Initialize
        tName1 = "dummy"
        tName2 = "dummy"
End Function
