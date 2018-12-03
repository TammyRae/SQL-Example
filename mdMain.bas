Attribute VB_Name = "mdMain"
Option Explicit
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, _
 ByVal lpOperation As String, _
 ByVal lpFile As String, _
 ByVal lpParameters As String, _
 ByVal lpDirectory As String, _
 ByVal nShowCmd As Long) _
 As Long
 
  


Rem ****************************************************************************
Rem *
Rem * Function Name:    RunReport
Rem * Paramters    :    acUserName      - User name to login to cad database
Rem *                   acPassword      - Password to use to login to database.
Rem *                   acDBInstance    - The CAD db instance
Rem *                   acOutputName    - File name to same output (spreadsheet) to.
Rem *
Rem * Description   : Main program which will extract the information from the
Rem *                 database and produce a report in excel format.
Rem ****************************************************************************
Sub RunReport(acUserName As String, acPassword As String, _
              acDBInstance As String, acExcelFile As String, acReportType As String)
  
  
    Select Case frmMain.lstCategory.ListIndex
    
        Case 0, 1   '* 0 = SOCC Unit Activity, 1 = SOCC Unit Activity with Crew Totals
            GenerateUnitActivityReport acUserName, acPassword, acDBInstance, acExcelFile, acReportType
            
'PAS Commented Out due to adjustment made to index, and removel of a report from list
'       Case 2      '** 3 = SOCC Unit Activity NEW (with Service Center data)
'       GenerateNewUnitActivityReport acUserName, acPassword, acDBInstance, acExcelFile, acReportType
            
        Case 2     '** 3 = StreetLights Defects
            GenerateStreetLightsReport acUserName, acPassword, acDBInstance, acExcelFile, acReportType
        
        Case 3    '** 3 = SteetLights 48hr Defects
            Generate48HourStreetLightsReport acUserName, acPassword, acDBInstance, acExcelFile, acReportType
            
'        Case 4     '** 3 = ETR Data from Neural Net
'            GenerateETRData "usercad", "test", "cadv8", acExcelFile, acReportType
            
        Case 4
            GenerateWaterUnitActivityReport acUserName, acPassword, acDBInstance, acExcelFile, acReportType
            
        Case 5
            GenerateTreeUnitActivityReport acUserName, acPassword, acDBInstance, acExcelFile, acReportType
            
        Case 6
            GenerateNewUnitActivityReportWater acUserName, acPassword, acDBInstance, acExcelFile, acReportType
            
        Case 7
            GenerateAllOutages acUserName, acPassword, acDBInstance, acExcelFile, acReportType
            
        Case 8
            GenerateSOCCPriorityTicketsSummary acUserName, acPassword, acDBInstance, acExcelFile, acReportType
            
        Case 9
            GenerateSOCCRoutineTicketsSummary acUserName, acPassword, acDBInstance, acExcelFile, acReportType
    
        Case 10
            GenerateWaterPriorityTicketsSummary acUserName, acPassword, acDBInstance, acExcelFile, acReportType
            
        Case 11
            GenerateWaterRoutineTicketsSummary acUserName, acPassword, acDBInstance, acExcelFile, acReportType
            
        Case 12
            GenerateCircuitSwitches "FMS_RPT", "pr4OMSrpt", "OMPR11.JEA.COM", acExcelFile, acReportType
            
        Case 13
            GenerateOutagesByDevice "FMS_RPT", "pr4OMSrpt", "OMPR11.JEA.COM", acExcelFile, acReportType
            
        Case 14
            GenerateServiceCenterUnitActivityReport acUserName, acPassword, acDBInstance, acExcelFile, acReportType
    End Select
    
   
    
End Sub






'**************************************************************************
' Excel Stuff
'**************************************************************************

Sub StartExcel(excelApp As Excel.Application)
On Error GoTo err:

Set excelApp = GetObject(, "Excel.Application") ' Create Excel Object.


Exit Sub
err:
Set excelApp = CreateObject("Excel.Application") 'Create Excel Object.

End Sub






