Attribute VB_Name = "mdSOCCReferredTicketsSummary"
Option Explicit

'******************************************
'*** Excel column constants for report ****
'******************************************
Public Const PRIORITY_1_LIST As String = "'FBLD', 'INIT', 'LBD', 'LO', 'LPAR', 'META', 'PBRO', 'PDOW', 'PFIR',' PHBC', 'POLB', 'TDWP', 'TXFP', 'TXHB', 'TXOL', 'TXFG', 'UCCC', 'UCEX', 'UTUW', 'WAAH', 'WAAP', 'WAIG', 'WAIT', 'WBAH', 'WBAP', 'WBAT', 'WDLP'"

Public Const A_CAD_ID_COL As Integer = 1
Public Const A_FINAL_SERVICE_CODE_COL As Integer = 2
Public Const A_JOB_STATUS_ACTIVE As Integer = 3
Public Const A_PRIORITY_FLAG_COL As Integer = 4
Public Const A_CIRCUIT_COL As Integer = 5
Public Const A_ZONE_COL As Integer = 6
Public Const A_REPORT_TOTAL_COL As Integer = 7
Public Const A_FIRST_CALL_COL As Integer = 8
'Public Const A_CAD_ID_CREATED_COL As Integer = 8
Public Const A_ENERGIZED_DATE_COL As Integer = 9
Public Const A_CREATE_TO_RESTORE_COL As Integer = 10
Public Const A_RESTORED_BY_COL As Integer = 11
Public Const A_SOCC_ASSIGN_COL As Integer = 12
Public Const A_SOCC_FIRST_CREW_COL As Integer = 13
Public Const A_SOCC_ARRIVED_COL As Integer = 14
Public Const A_SOCC_REPORTING_COL As Integer = 15
'Public Const A_CREATE_TO_CAD_COL As Integer = 16
Public Const A_SOCC_CREATE_TO_ARRIVE_COL As Integer = 16
Public Const A_SOCC_CREATE_TO_REPORT_COL As Integer = 17
Public Const A_SOCC_ARRIVED_TO_REPORT_COL As Integer = 18
Public Const A_SOCC_ASSIGN_TO_ARRIVE_COL As Integer = 19
Public Const A_SOCC_REFERED_FLAG_COL As Integer = 20
Public Const A_SOCC_REFERED_TO_COL As Integer = 21
Public Const A_SOCC_CREATE_TO_REFER_COL As Integer = 22
Public Const A_SOCC_REFERED_TIME_COL As Integer = 23
Public Const A_SC_PRIORITY_REFER_TIME_COL As Integer = 24
Public Const A_SC_REFER_FROM_COL As Integer = 25
Public Const A_SC_REFER_TO_COL As Integer = 26
Public Const A_SC_UNIT_COL As Integer = 27
Public Const A_SC_ASSIGN_COL As Integer = 28
Public Const A_SC_ARRIVE_COL As Integer = 29
Public Const A_SC_REPORT_COL As Integer = 30
Public Const A_SC_CREATE_TO_ARRIVE_COL As Integer = 31
Public Const A_SC_CREATE_TO_REPORT_COL As Integer = 32
Public Const A_SC_REFER_TO_ARRIVE_COL = 33
Public Const A_SC_REFER_TO_REPORT_COL = 34
Public Const A_SC_ARRIVE_TO_REPORT_COL = 35
Public Const A_TREE_REFER_TIME_COL As Integer = 36
Public Const A_TREE_REFER_TO_COL As Integer = 37
Public Const A_TREE_UNIT_COL As Integer = 38
Public Const A_TREE_ASSIGN_COL As Integer = 39
Public Const A_TREE_ARRIVE_COL As Integer = 40
Public Const A_TREE_REPORT_COL As Integer = 41
Public Const A_TREE_CREATE_TO_ARRIVE_COL As Integer = 42
Public Const A_TREE_CREATE_TO_REPORT_COL As Integer = 43
Public Const A_TREE_REFER_TO_ARRIVE_COL = 44
Public Const A_TREE_REFER_TO_REPORT_COL = 45
Public Const A_TREE_ARRIVE_TO_REPORT_COL = 46
Public Const A_JOB_DISPOSE_TIME_COL As Integer = 47
Public Const A_TOTAL_CALLS_COL As Integer = 48
Public Const A_TOTAL_CMO_COL As Integer = 49
Public Const A_TOTAL_CUSTOMERS_COL As Integer = 50


Rem ****************************************************************************
Rem *
Rem * Function Name:    GenerateSOCCRoutineTicketsSummary
Rem * Paramters    :    acUserName      - User name to login to cad database
Rem *                   acPassword      - Password to use to login to database.
Rem *                   acDBInstance    - The CAD db instance
Rem *                   acOutputName    - File name to same output (spreadsheet) to.
Rem *
Rem * Description   : Generate the unit activity report
Rem ****************************************************************************
Public Sub GenerateSOCCRoutineTicketsSummary(acUserName As String, acPassword As String, _
                                       acDBInstance As String, acExcelFile As String, acReportType As String)
    Dim PriorityQuery As String
    
    PriorityQuery = " select distinct isr_no from r01_job_report "
    PriorityQuery = PriorityQuery + " where r01_create_date >= '" + Format(frmMain.lstDatePickerFrom.Value, "YYYYMMDD") + "' "
    PriorityQuery = PriorityQuery + " and r01_create_date <= '" + Format(frmMain.lstDatePickerTo.Value, "YYYYMMDD") + "' "
    PriorityQuery = PriorityQuery + " and isr_no like 'EL%' and service_code not in (" + PRIORITY_1_LIST + ") "
    
    MainSpreadSheetGeneration acUserName, acPassword, acDBInstance, acExcelFile, acReportType, PriorityQuery, "SOCC Routine Summary"

End Sub
Rem ****************************************************************************
Rem *
Rem * Function Name:    GenerateSOCCPriorityTicketsSummary
Rem * Paramters    :    acUserName      - User name to login to cad database
Rem *                   acPassword      - Password to use to login to database.
Rem *                   acDBInstance    - The CAD db instance
Rem *                   acOutputName    - File name to same output (spreadsheet) to.
Rem *
Rem * Description   : Generate the unit activity report
Rem ****************************************************************************
Public Sub GenerateSOCCPriorityTicketsSummary(acUserName As String, acPassword As String, _
                                       acDBInstance As String, acExcelFile As String, acReportType As String)
    Dim PriorityQuery As String
    
    PriorityQuery = " select distinct isr_no from r01_job_report "
    PriorityQuery = PriorityQuery + " where r01_create_date >= '" + Format(frmMain.lstDatePickerFrom.Value, "YYYYMMDD") + "' "
    PriorityQuery = PriorityQuery + " and r01_create_date <= '" + Format(frmMain.lstDatePickerTo.Value, "YYYYMMDD") + "' "
    PriorityQuery = PriorityQuery + " and isr_no like 'EL%' and service_code in (" + PRIORITY_1_LIST + ") "
    
    MainSpreadSheetGeneration acUserName, acPassword, acDBInstance, acExcelFile, acReportType, PriorityQuery, "SOCC Priority Summary"

End Sub

Rem ****************************************************************************
Rem *
Rem * Function Name:    MainSpreadSheetGeneration
Rem * Paramters    :    acUserName      - User name to login to cad database
Rem *                   acPassword      - Password to use to login to database.
Rem *                   acDBInstance    - The CAD db instance
Rem *                   acOutputName    - File name to same output (spreadsheet) to.
Rem *                   acMainQuery     - Query to use as the main query.
Rem *                   acReportTitle   - Title to set in the spreadsheet tab.
Rem *
Rem * Description   : Generate the unit activity report
Rem ****************************************************************************
Private Sub MainSpreadSheetGeneration(acUserName As String, acPassword As String, _
                                      acDBInstance As String, acExcelFile As String, acReportType As String, _
                                      acMainQuery As String, acReportTitle As String)
    Dim acSQLText As String, acIsrNo As String, acUnitNo As String
    Dim acCellData As String
    Dim acSOCCUnit As String, acServiceCenterUnit As String, acTreeUnit As String
    Dim acReferToServiceCenterTime As String, acReferToTreeTime As String
    Dim acCreatedTime As String, blPriority1 As Boolean
    Dim acProblemCode As String, iPriorityPosition As Integer, iPriority2Position As Integer
    Dim i As Integer, x As Integer, iTotalRows As Double, iRowPosition As Double
    Dim iCrewTotal As Integer, iTimeDifference As Integer
    Dim iCounter As Integer, acReferUnit As String
    Dim excelExcelApp As Excel.Application ' This is the excel program
    Dim excelWorkBook As Excel.Workbook ' This is the work book
    Dim excelWorksheet As Excel.Worksheet ' This is the sheet
    Dim adoConnection As ADODB.Connection
    Dim adoIsrList As ADODB.Recordset
    Dim adoIsrDetail As ADODB.Recordset
    Dim adoIsrDetail2 As ADODB.Recordset
    Dim adoIsrDetail3 As ADODB.Recordset
    Dim adoIsrDetail4 As ADODB.Recordset
    Dim adoIsrDetail5 As ADODB.Recordset
    Dim adoIsrDetail6 As ADODB.Recordset
    Dim adoIsrDetail7 As ADODB.Recordset
    Dim adoIsrDetail8 As ADODB.Recordset
    Dim adoIsrDetail9 As ADODB.Recordset
    Dim adoIsrDispose As ADODB.Recordset
    Dim adoJobReferal As ADODB.Recordset
    
    Set adoConnection = New ADODB.Connection
    Set adoIsrList = New ADODB.Recordset
    Set adoIsrDetail = New ADODB.Recordset
    Set adoIsrDetail2 = New ADODB.Recordset
    Set adoIsrDetail3 = New ADODB.Recordset
    Set adoIsrDetail4 = New ADODB.Recordset
    Set adoIsrDetail5 = New ADODB.Recordset
    Set adoIsrDetail6 = New ADODB.Recordset
    Set adoIsrDetail7 = New ADODB.Recordset
    Set adoIsrDetail8 = New ADODB.Recordset
    Set adoIsrDetail9 = New ADODB.Recordset
    Set adoIsrDispose = New ADODB.Recordset
    Set adoJobReferal = New ADODB.Recordset

    
    'connect to the Oracle database
    frmMain.MousePointer = vbHourglass
    frmMain.sbStatusBar.Panels(1).Text = "Opening connection to database..."
    frmMain.cmdGoButton.Enabled = False
    frmMain.cmdClose.Enabled = False
    adoConnection.Open "Provider=MSDASQL;Persist Security Info=False;User ID=" & acUserName & ";pwd=" & acPassword & ";Data Source=" & acDBInstance & ";"
    adoConnection.CursorLocation = adUseClient
    adoConnection.CommandTimeout = 60

    ' If no directory is specified, put in current directory
    If (InStr(acExcelFile, "\") = 0) Then
        acExcelFile = CurDir() + "\" + Trim(acExcelFile)
    End If
    
    StartExcel excelExcelApp
    Set excelWorkBook = excelExcelApp.Workbooks.Add
    
    acSQLText = acMainQuery

    frmMain.sbStatusBar.Panels(1).Text = "Running query..."
    adoIsrList.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText

    If Not (adoIsrList.BOF And adoIsrList.EOF) Then
        iTotalRows = adoIsrList.RecordCount
          
        If iTotalRows > 0 Then
            'For iCounter = excelWorkBook.Worksheets.Count To 4  '** There is usually 3 worksheets open by default in excel
            '    excelWorkBook.Worksheets.Add
            'Next
            
            ' Create each worksheet and the column headers.
            Set excelWorksheet = excelWorkBook.Worksheets(1)
            excelWorksheet.Name = acReportTitle
            excelWorksheet.Select
            CreateWorksheetHeader excelWorksheet, excelWorkBook, excelExcelApp, 1
            'Set excelWorksheet = excelWorkBook.Worksheets(2)
            'excelWorksheet.Name = "TREE-PRIORITY"
            'excelWorksheet.Select
            'CreateWorksheetHeader excelWorksheet, excelWorkBook, excelExcelApp, 1
            'Set excelWorksheet = excelWorkBook.Worksheets(3)
            'excelWorksheet.Name = "3PRT"
            'excelWorksheet.Select
            'CreateWorksheetHeader excelWorksheet, excelWorkBook, excelExcelApp, 1
            'Set excelWorksheet = excelWorkBook.Worksheets(4)
            'excelWorksheet.Select
            'excelWorksheet.Name = "SOCC"
            'CreateWorksheetHeader excelWorksheet, excelWorkBook, excelExcelApp, 1
            'Set excelWorksheet = excelWorkBook.Worksheets(5)
            'excelWorksheet.Select
            'excelWorksheet.Name = "SOCC-QRPT"
            'CreateWorksheetHeader excelWorksheet, excelWorkBook, excelExcelApp, 1
           
            'Update the progress bar
            frmMain.ProgressBar1.Max = iTotalRows + 1
            frmMain.ProgressBar1.Min = 1
            frmMain.ProgressBar1.Visible = True
            
            i = 0
            iPriorityPosition = 2
            iPriority2Position = 2
            iRowPosition = 2
            adoIsrList.MoveFirst
            acIsrNo = adoIsrList.Fields("isr_no")
            
            Set excelWorksheet = excelWorkBook.Worksheets(1)
            Do While Not adoIsrList.EOF
                acIsrNo = adoIsrList.Fields.Item("ISR_NO")
                
                'excelWorksheet.Cells(iRowPosition, A_CAD_ID_COL) = acIsrNo
                #If EnvironmentDEV Then
                    excelWorksheet.Hyperlinks.Add Anchor:=excelWorksheet.Cells(iRowPosition, A_CAD_ID_COL), Address:="http://fmsoms-dv/fmsweb/CadDetail.aspx?ID=" + acIsrNo, TextToDisplay:=acIsrNo
                #End If
                #If EnvironmentQA Then
                    excelWorksheet.Hyperlinks.Add Anchor:=excelWorksheet.Cells(iRowPosition, A_CAD_ID_COL), Address:="http://fmsoms-qa/fmsweb/CadDetail.aspx?ID=" + acIsrNo, TextToDisplay:=acIsrNo
                #End If
                #If EnvironmentPR Then
                    excelWorksheet.Hyperlinks.Add Anchor:=excelWorksheet.Cells(iRowPosition, A_CAD_ID_COL), Address:="http://fmsweb-act/fmsweb/CadDetail.aspx?ID=" + acIsrNo, TextToDisplay:=acIsrNo
                #End If
                
                excelWorksheet.Cells(iRowPosition, A_PRIORITY_FLAG_COL) = "Y"
                
               
                '*********************************************
                '* Get the isr information
                '*********************************************
                acSQLText = "select final_service_code, dispatch_level, s_mode, "
                acSQLText = acSQLText + "to_date(created_date || created_time, 'YYYYMMDDHH24MISS') creation_date "
                acSQLText = acSQLText + "from isr where isr_no = '" + Trim(acIsrNo) + "' "
                adoIsrDetail.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                
                If Not (adoIsrDetail.BOF And adoIsrDetail.EOF) Then
                    excelWorksheet.Cells(iRowPosition, A_FINAL_SERVICE_CODE_COL) = adoIsrDetail.Fields("final_service_code")
                    excelWorksheet.Cells(iRowPosition, A_ZONE_COL) = adoIsrDetail.Fields("dispatch_level")
                    Select Case adoIsrDetail.Fields("s_mode").Value
                    Case "A"
                        excelWorksheet.Cells(iRowPosition, A_JOB_STATUS_ACTIVE) = "Active"
                    Case "L"
                        excelWorksheet.Cells(iRowPosition, A_JOB_STATUS_ACTIVE) = "Disposed"
                    Case "D"
                        excelWorksheet.Cells(iRowPosition, A_JOB_STATUS_ACTIVE) = "Differed"
                    Case Else
                        excelWorksheet.Cells(iRowPosition, A_JOB_STATUS_ACTIVE) = adoIsrDetail.Fields("s_mode").Value
                    End Select
'                   excelWorksheet.Cells(iRowPosition, A_CAD_ID_CREATED_COL) = adoIsrDetail.Fields("creation_date")
                End If
                adoIsrDetail.Close
                    
                '*********************************************
                '* Get the dispose information
                '*********************************************
                acSQLText = "select to_date(max(mod_date || mod_time), 'YYYYMMDDHH24MISS') dispose_time "
                acSQLText = acSQLText + "from c02_isr_dispos where isr_no = '" + Trim(acIsrNo) + "'"
                adoIsrDispose.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                
                If Not (adoIsrDispose.BOF And adoIsrDispose.EOF) Then
                    excelWorksheet.Cells(iRowPosition, A_JOB_DISPOSE_TIME_COL) = adoIsrDispose.Fields("dispose_time")
                End If
                adoIsrDispose.Close
                
                
                '*********************************************
                '* Get the referal data for SOCC
                '*********************************************
                acSQLText = "select fieldvalue refer, to_date(mod_date || mod_time, 'YYYYMMDDHH24MISS') refer_date "
                acSQLText = acSQLText + "from isr_log "
                acSQLText = acSQLText + "where isr_no = '" + Trim(acIsrNo) + "' AND FIELDNAME = 'p_filter1'"
                acSQLText = acSQLText + "and originalvalue = 'SOCC' "
                adoJobReferal.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                
                If Not (adoJobReferal.BOF And adoJobReferal.EOF) Then
                    excelWorksheet.Cells(iRowPosition, A_SOCC_REFERED_FLAG_COL) = "Y"
                    excelWorksheet.Cells(iRowPosition, A_SOCC_REFERED_TO_COL) = adoJobReferal.Fields("refer")
                    excelWorksheet.Cells(iRowPosition, A_SOCC_REFERED_TIME_COL) = adoJobReferal.Fields("refer_date")
                Else
                    excelWorksheet.Cells(iRowPosition, A_SOCC_REFERED_FLAG_COL) = "N"
                End If
                adoJobReferal.Close
                
                '*********************************************
                '* Get the referal data for Service Center
                '*********************************************
                acSQLText = "select originalvalue, fieldvalue refer, to_date(mod_date || mod_time, 'YYYYMMDDHH24MISS') refer_date, "
                acSQLText = acSQLText + " mod_date || mod_time refer_date_flat_format "
                acSQLText = acSQLText + "From isr_log "
                acSQLText = acSQLText + "where isr_no = '" + Trim(acIsrNo) + "' AND FIELDNAME = 'p_filter1' "
                acSQLText = acSQLText + "and originalvalue not in ('3-PRT', 'SS-PRIORTY', 'SS-ROUTINE', 'SSSC-ENG', 'SSSC-XP', "
                acSQLText = acSQLText + "'WS-PRIORTY', 'WS-ROUTINE', 'WSSC-ENG', 'WSSC-XP') "
                acSQLText = acSQLText + "and fieldvalue in ('3-PRT', 'SS-PRIORTY', 'SS-ROUTINE', 'SSSC-ENG', 'SSSC-XP', "
                acSQLText = acSQLText + "'WS-PRIORTY', 'WS-ROUTINE', 'WSSC-ENG', 'WSSC-XP') "
                adoJobReferal.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                
                acReferToServiceCenterTime = ""
                If Not (adoJobReferal.BOF And adoJobReferal.EOF) Then
                    excelWorksheet.Cells(iRowPosition, A_SC_REFER_FROM_COL) = adoJobReferal("originalvalue")
                    excelWorksheet.Cells(iRowPosition, A_SC_REFER_TO_COL) = adoJobReferal.Fields("refer")
                    excelWorksheet.Cells(iRowPosition, A_SC_PRIORITY_REFER_TIME_COL) = adoJobReferal.Fields("refer_date")
                    
                    acReferToServiceCenterTime = adoJobReferal.Fields("refer_date_flat_format")
                End If
                adoJobReferal.Close
                
                '*********************************************
                '* Get the referal data for TREE
                '*********************************************
                acSQLText = "select fieldvalue refer, to_date(mod_date || mod_time, 'YYYYMMDDHH24MISS') refer_date, "
                acSQLText = acSQLText + " mod_date || mod_time refer_date_flat_format "
                acSQLText = acSQLText + "From isr_log "
                acSQLText = acSQLText + "where isr_no = '" + Trim(acIsrNo) + "' AND FIELDNAME = 'p_filter1' "
                acSQLText = acSQLText + "and originalvalue not in ('T- PRIORTY', 'T -ROUTINE', 'T -STORM') "
                acSQLText = acSQLText + "and fieldvalue in ('T- PRIORTY', 'T -ROUTINE', 'T -STORM') "
                adoJobReferal.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                
                acReferToTreeTime = ""
                If Not (adoJobReferal.BOF And adoJobReferal.EOF) Then
                    excelWorksheet.Cells(iRowPosition, A_TREE_REFER_TO_COL) = adoJobReferal.Fields("refer")
                    excelWorksheet.Cells(iRowPosition, A_TREE_REFER_TIME_COL) = adoJobReferal.Fields("refer_date")
                    
                    acReferToTreeTime = adoJobReferal.Fields("refer_date_flat_format")
                End If
                adoJobReferal.Close
                
                '*********************************************
                '* Get first crew assign date/time
                '*********************************************
                acSQLText = "select unit, to_date(d11_start_date || d11_start_time, 'YYYYMMDDHH24MISS') assign_date "
                acSQLText = acSQLText + "from d11_unit_activity where isr_no = '" + Trim(acIsrNo) + "' "
                acSQLText = acSQLText + "and b07_status_code = 'DP' "
                acSQLText = acSQLText + "and (d11_start_date || d11_start_time) in "
                acSQLText = acSQLText + "(select min(d11_start_date || d11_start_time) "
                acSQLText = acSQLText + "from d11_unit_activity where isr_no = '" + Trim(acIsrNo) + "' "
                acSQLText = acSQLText + "and b07_status_code = 'DP') "
                adoIsrDetail2.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                
                If Not (adoIsrDetail2.BOF And adoIsrDetail2.EOF) Then
                    excelWorksheet.Cells(iRowPosition, A_SOCC_ASSIGN_COL) = adoIsrDetail2.Fields("assign_date")
                    excelWorksheet.Cells(iRowPosition, A_SOCC_FIRST_CREW_COL) = adoIsrDetail2.Fields("unit")
                End If
                adoIsrDetail2.Close
                
                '*********************************************
                '* Get first field report (SOCC crew reporting date)
                '*********************************************
                acSQLText = "select unit, to_date(r01_create_date || r01_create_time, 'YYYYMMDDHH24MISS') report_date "
                acSQLText = acSQLText + "from r01_job_report where isr_no = '" + Trim(acIsrNo) + "' "
                acSQLText = acSQLText + "and r01_create_date || r01_create_time = (select min(r01_create_date || r01_create_time) "
                acSQLText = acSQLText + "    from r01_job_report where isr_no = '" + Trim(acIsrNo) + "') "
                adoIsrDetail3.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                
                If Not (adoIsrDetail3.BOF And adoIsrDetail3.EOF) Then
                    excelWorksheet.Cells(iRowPosition, A_SOCC_REPORTING_COL) = adoIsrDetail3.Fields("report_date")
                    acSOCCUnit = adoIsrDetail3.Fields("unit")
                Else
                    acSOCCUnit = ""
                End If
                adoIsrDetail3.Close
                
                '*********************************************
                '* Get the SOCC arrived time
                '*********************************************
                If Trim(acSOCCUnit) <> "" Then
                    acSQLText = "select to_date(max(d11_start_date || d11_start_time), 'YYYYMMDDHH24MISS') arrived "
                    acSQLText = acSQLText + "From d11_unit_activity "
                    acSQLText = acSQLText + "where isr_no = '" + Trim(acIsrNo) + "' "
                    acSQLText = acSQLText + "and unit = '" + acSOCCUnit + "' "
                    acSQLText = acSQLText + "and b07_status_code = 'AR' "
                    adoIsrDetail4.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                    
                    If Not (adoIsrDetail4.BOF And adoIsrDetail4.EOF) Then
                        excelWorksheet.Cells(iRowPosition, A_SOCC_ARRIVED_COL) = adoIsrDetail4.Fields("arrived")
                    End If
                    adoIsrDetail4.Close
                End If
                
                '*********************************************
                '* Get the total reports for this job.
                '*********************************************
                acSQLText = "select count(*) total_reports from r01_job_report where isR_no = '" + Trim(acIsrNo) + "' "
                adoIsrDetail5.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                If Not (adoIsrDetail5.BOF And adoIsrDetail5.EOF) Then
                    excelWorksheet.Cells(iRowPosition, A_REPORT_TOTAL_COL) = adoIsrDetail5.Fields("total_reports")
                End If
                adoIsrDetail5.Close
                
                
                '*********************************************
                '* Get the outage information for this job.
                '*********************************************
                acSQLText = "select cad_id, hl.creation_datetime, hl.energized_datetime, hi.circt_id, "
                acSQLText = acSQLText + "hi.equip_stn_no, hi.dni_equip_type, "
                acSQLText = acSQLText + "hi.downstream_cust_qty , hi.call_qty "
                acSQLText = acSQLText + "from his_location@fmsrpt_ompr11_l hl, his_incident_device@fmsrpt_ompr11_l hi "
                acSQLText = acSQLText + "Where hl.incident_device_id = hi.incident_device_id "
                acSQLText = acSQLText + "AND CAD_ID = '" + acIsrNo + "' "
                adoIsrDetail6.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                If adoIsrDetail6.BOF Or adoIsrDetail6.EOF Then
                    '* Look into the active tables.  Job may not been archived yet.
                    adoIsrDetail6.Close
                    acSQLText = "select cad_id, hl.creation_datetime, hl.energized_datetime, hi.circt_id, "
                    acSQLText = acSQLText + "hi.equip_stn_no, hi.dni_equip_type, "
                    acSQLText = acSQLText + "hi.downstream_cust_qty , hi.call_qty "
                    acSQLText = acSQLText + "from location@fmsrpt_ompr11_l hl, incident_device@fmsrpt_ompr11_l hi "
                    acSQLText = acSQLText + "Where hl.incident_device_id = hi.incident_device_id "
                    acSQLText = acSQLText + "AND CAD_ID = '" + acIsrNo + "' "
                    adoIsrDetail6.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                End If
                If Not (adoIsrDetail6.BOF And adoIsrDetail6.EOF) Then
                    excelWorksheet.Cells(iRowPosition, A_CIRCUIT_COL) = adoIsrDetail6.Fields("circt_id")
                    excelWorksheet.Cells(iRowPosition, A_FIRST_CALL_COL) = adoIsrDetail6.Fields("creation_datetime")
                    excelWorksheet.Cells(iRowPosition, A_ENERGIZED_DATE_COL) = adoIsrDetail6.Fields("energized_datetime")
                    excelWorksheet.Cells(iRowPosition, A_TOTAL_CALLS_COL) = adoIsrDetail6.Fields("call_qty")
                    excelWorksheet.Cells(iRowPosition, A_TOTAL_CUSTOMERS_COL) = adoIsrDetail6.Fields("downstream_cust_qty")
                    
                    excelWorksheet.Cells(iRowPosition, A_TOTAL_CMO_COL) = "=+((RC[-" + Trim(Str(A_TOTAL_CMO_COL - A_ENERGIZED_DATE_COL)) + "]-RC[-" + Trim(Str(A_TOTAL_CMO_COL - A_FIRST_CALL_COL)) + "]) * 1440) * RC[" + Trim(Str(A_TOTAL_CUSTOMERS_COL - A_TOTAL_CMO_COL)) + "]"
                End If
                adoIsrDetail6.Close

                
                '*********************************************
                '*********************************************
                '* Get the Service Center crew data
                '*********************************************
                '*********************************************
                
                If Trim(acReferToServiceCenterTime) <> "" Then
                    '*********************************************
                    '* Get first field report after referal from SOCC
                    '*********************************************
                    acSQLText = "select unit, to_date(r01_create_date || r01_create_time, 'YYYYMMDDHH24MISS') report_date "
                    acSQLText = acSQLText + "from r01_job_report where isr_no = '" + Trim(acIsrNo) + "' "
                    acSQLText = acSQLText + "and r01_create_date || r01_create_time = (select min(r01_create_date || r01_create_time) "
                    acSQLText = acSQLText + "    from r01_job_report where isr_no = '" + Trim(acIsrNo) + "' and r01_create_date || r01_create_time > '" + acReferToServiceCenterTime + "') "
                    adoIsrDetail7.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                    
                    If Not (adoIsrDetail7.BOF And adoIsrDetail7.EOF) Then
                        excelWorksheet.Cells(iRowPosition, A_SC_REPORT_COL) = adoIsrDetail7.Fields("report_date")
                        excelWorksheet.Cells(iRowPosition, A_SC_UNIT_COL) = adoIsrDetail7.Fields("unit")
                        acServiceCenterUnit = adoIsrDetail7.Fields("unit")
                    Else
                        acServiceCenterUnit = ""
                    End If
                    adoIsrDetail7.Close
                                
                    '***************************************************
                    '* Get the Service Center assigned and arrived time
                    '***************************************************
                    If Trim(acServiceCenterUnit) <> "" Then
                        acSQLText = "select to_date(max(d11_start_date || d11_start_time), 'YYYYMMDDHH24MISS') arrived "
                        acSQLText = acSQLText + "From d11_unit_activity "
                        acSQLText = acSQLText + "where isr_no = '" + Trim(acIsrNo) + "' "
                        acSQLText = acSQLText + "and unit = '" + acServiceCenterUnit + "' "
                        acSQLText = acSQLText + "and b07_status_code = 'DP' "
                        adoIsrDetail8.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                        
                        If Not (adoIsrDetail8.BOF And adoIsrDetail8.EOF) Then
                            excelWorksheet.Cells(iRowPosition, A_SC_ASSIGN_COL) = adoIsrDetail8.Fields("arrived")
                        End If
                        adoIsrDetail8.Close
                        
                        acSQLText = "select to_date(max(d11_start_date || d11_start_time), 'YYYYMMDDHH24MISS') arrived "
                        acSQLText = acSQLText + "From d11_unit_activity "
                        acSQLText = acSQLText + "where isr_no = '" + Trim(acIsrNo) + "' "
                        acSQLText = acSQLText + "and unit = '" + acServiceCenterUnit + "' "
                        acSQLText = acSQLText + "and b07_status_code = 'AR' "
                        adoIsrDetail9.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                        
                        If Not (adoIsrDetail9.BOF And adoIsrDetail9.EOF) Then
                            excelWorksheet.Cells(iRowPosition, A_SC_ARRIVE_COL) = adoIsrDetail9.Fields("arrived")
                        End If
                        adoIsrDetail9.Close
                        
                        '* Derived columns for Service Center
                        excelWorksheet.Cells(iRowPosition, A_SC_CREATE_TO_ARRIVE_COL) = "=+(RC[-" + Trim(Str(A_SC_CREATE_TO_ARRIVE_COL - A_SC_ARRIVE_COL)) + "]-RC[-" + Trim(Str(A_SC_CREATE_TO_ARRIVE_COL - A_FIRST_CALL_COL)) + "]) * 1440"
                        excelWorksheet.Cells(iRowPosition, A_SC_CREATE_TO_REPORT_COL) = "=+(RC[-" + Trim(Str(A_SC_CREATE_TO_REPORT_COL - A_SC_REPORT_COL)) + "]-RC[-" + Trim(Str(A_SC_CREATE_TO_REPORT_COL - A_FIRST_CALL_COL)) + "]) * 1440"
                        excelWorksheet.Cells(iRowPosition, A_SC_ARRIVE_TO_REPORT_COL) = "=+(RC[-" + Trim(Str(A_SC_ARRIVE_TO_REPORT_COL - A_SC_REPORT_COL)) + "]-RC[-" + Trim(Str(A_SC_ARRIVE_TO_REPORT_COL - A_SC_ARRIVE_COL)) + "]) * 1440"
                        excelWorksheet.Cells(iRowPosition, A_SC_REFER_TO_ARRIVE_COL) = "=+(RC[-" + Trim(Str(A_SC_REFER_TO_ARRIVE_COL - A_SC_ARRIVE_COL)) + "]-RC[-" + Trim(Str(A_SC_REFER_TO_ARRIVE_COL - A_SC_PRIORITY_REFER_TIME_COL)) + "]) * 1440"
                        excelWorksheet.Cells(iRowPosition, A_SC_REFER_TO_REPORT_COL) = "=+(RC[-" + Trim(Str(A_SC_REFER_TO_REPORT_COL - A_SC_REPORT_COL)) + "]-RC[-" + Trim(Str(A_SC_REFER_TO_REPORT_COL - A_SC_PRIORITY_REFER_TIME_COL)) + "]) * 1440"
                        
                    End If
                End If
                
                
                '*********************************************
                '*********************************************
                '* Get the Tree Crew crew data
                '*********************************************
                '*********************************************
                
                If Trim(acReferToTreeTime) <> "" Then
                    '*********************************************
                    '* Get first field report after referal to tree
                    '*********************************************
                    acSQLText = "select unit, to_date(r01_create_date || r01_create_time, 'YYYYMMDDHH24MISS') report_date "
                    acSQLText = acSQLText + "from r01_job_report where isr_no = '" + Trim(acIsrNo) + "' "
                    acSQLText = acSQLText + "and r01_create_date || r01_create_time = (select min(r01_create_date || r01_create_time) "
                    acSQLText = acSQLText + "    from r01_job_report where isr_no = '" + Trim(acIsrNo) + "' and r01_create_date || r01_create_time > '" + acReferToTreeTime + "') "
                    adoIsrDetail.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                    
                    If Not (adoIsrDetail.BOF And adoIsrDetail.EOF) Then
                        excelWorksheet.Cells(iRowPosition, A_TREE_REPORT_COL) = adoIsrDetail.Fields("report_date")
                        excelWorksheet.Cells(iRowPosition, A_TREE_UNIT_COL) = adoIsrDetail.Fields("unit")
                        acTreeUnit = adoIsrDetail.Fields("unit")
                    Else
                        acTreeUnit = ""
                    End If
                    adoIsrDetail.Close
                                
                    '***************************************************
                    '* Get the Tree crew assigned and arrived time
                    '***************************************************
                    If Trim(acTreeUnit) <> "" Then
                        acSQLText = "select to_date(max(d11_start_date || d11_start_time), 'YYYYMMDDHH24MISS') arrived "
                        acSQLText = acSQLText + "From d11_unit_activity "
                        acSQLText = acSQLText + "where isr_no = '" + Trim(acIsrNo) + "' "
                        acSQLText = acSQLText + "and unit = '" + acTreeUnit + "' "
                        acSQLText = acSQLText + "and b07_status_code = 'DP' "
                        adoIsrDetail.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                        
                        If Not (adoIsrDetail.BOF And adoIsrDetail.EOF) Then
                            excelWorksheet.Cells(iRowPosition, A_TREE_ASSIGN_COL) = adoIsrDetail.Fields("arrived")
                        End If
                        adoIsrDetail.Close
                        
                        acSQLText = "select to_date(max(d11_start_date || d11_start_time), 'YYYYMMDDHH24MISS') arrived "
                        acSQLText = acSQLText + "From d11_unit_activity "
                        acSQLText = acSQLText + "where isr_no = '" + Trim(acIsrNo) + "' "
                        acSQLText = acSQLText + "and unit = '" + acTreeUnit + "' "
                        acSQLText = acSQLText + "and b07_status_code = 'AR' "
                        adoIsrDetail.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                        
                        If Not (adoIsrDetail.BOF And adoIsrDetail.EOF) Then
                            excelWorksheet.Cells(iRowPosition, A_TREE_ARRIVE_COL) = adoIsrDetail.Fields("arrived")
                        End If
                        adoIsrDetail.Close
                        
                        '* Derived columns for Tree
                        excelWorksheet.Cells(iRowPosition, A_TREE_CREATE_TO_ARRIVE_COL) = "=+(RC[-" + Trim(Str(A_TREE_CREATE_TO_ARRIVE_COL - A_TREE_ARRIVE_COL)) + "]-RC[-" + Trim(Str(A_TREE_CREATE_TO_ARRIVE_COL - A_FIRST_CALL_COL)) + "]) * 1440"
                        excelWorksheet.Cells(iRowPosition, A_TREE_CREATE_TO_REPORT_COL) = "=+(RC[-" + Trim(Str(A_TREE_CREATE_TO_REPORT_COL - A_TREE_REPORT_COL)) + "]-RC[-" + Trim(Str(A_TREE_CREATE_TO_REPORT_COL - A_FIRST_CALL_COL)) + "]) * 1440"
                        excelWorksheet.Cells(iRowPosition, A_TREE_ARRIVE_TO_REPORT_COL) = "=+(RC[-" + Trim(Str(A_TREE_ARRIVE_TO_REPORT_COL - A_TREE_REPORT_COL)) + "]-RC[-" + Trim(Str(A_TREE_ARRIVE_TO_REPORT_COL - A_TREE_ARRIVE_COL)) + "]) * 1440"
                        excelWorksheet.Cells(iRowPosition, A_TREE_REFER_TO_ARRIVE_COL) = "=+(RC[-" + Trim(Str(A_TREE_REFER_TO_ARRIVE_COL - A_TREE_ARRIVE_COL)) + "]-RC[-" + Trim(Str(A_TREE_REFER_TO_ARRIVE_COL - A_TREE_REFER_TIME_COL)) + "]) * 1440"
                        excelWorksheet.Cells(iRowPosition, A_TREE_REFER_TO_REPORT_COL) = "=+(RC[-" + Trim(Str(A_TREE_REFER_TO_REPORT_COL - A_TREE_REPORT_COL)) + "]-RC[-" + Trim(Str(A_TREE_REFER_TO_REPORT_COL - A_TREE_REFER_TIME_COL)) + "]) * 1440"

                    End If
                End If

                '* Set the derived columns for SOCC (formulas in the spreadsheet)
                excelWorksheet.Cells(iRowPosition, A_CREATE_TO_RESTORE_COL) = "=+(RC[-" + Trim(Str(A_CREATE_TO_RESTORE_COL - A_ENERGIZED_DATE_COL)) + "]-RC[-" + Trim(Str(A_CREATE_TO_RESTORE_COL - A_FIRST_CALL_COL)) + "]) * 1440"
                'excelWorksheet.Cells(iRowPosition, A_CREATE_TO_CAD_COL) = "=+(RC[-" + Trim(Str(A_CREATE_TO_CAD_COL - A_CAD_ID_CREATED_COL)) + "]-RC[-" + Trim(Str(A_CREATE_TO_CAD_COL - A_FIRST_CALL_COL)) + "]) * 1440"
                excelWorksheet.Cells(iRowPosition, A_SOCC_CREATE_TO_ARRIVE_COL) = "=+(RC[-" + Trim(Str(A_SOCC_CREATE_TO_ARRIVE_COL - A_SOCC_ARRIVED_COL)) + "]-RC[-" + Trim(Str(A_SOCC_CREATE_TO_ARRIVE_COL - A_FIRST_CALL_COL)) + "]) * 1440"
                excelWorksheet.Cells(iRowPosition, A_SOCC_CREATE_TO_REPORT_COL) = "=+(RC[-" + Trim(Str(A_SOCC_CREATE_TO_REPORT_COL - A_SOCC_REPORTING_COL)) + "]-RC[-" + Trim(Str(A_SOCC_CREATE_TO_REPORT_COL - A_FIRST_CALL_COL)) + "]) * 1440"
                excelWorksheet.Cells(iRowPosition, A_SOCC_ARRIVED_TO_REPORT_COL) = "=+(RC[-" + Trim(Str(A_SOCC_ARRIVED_TO_REPORT_COL - A_SOCC_REPORTING_COL)) + "]-RC[-" + Trim(Str(A_SOCC_ARRIVED_TO_REPORT_COL - A_SOCC_ARRIVED_COL)) + "]) * 1440"
                excelWorksheet.Cells(iRowPosition, A_SOCC_ASSIGN_TO_ARRIVE_COL) = "=+(RC[-" + Trim(Str(A_SOCC_ASSIGN_TO_ARRIVE_COL - A_SOCC_ARRIVED_COL)) + "]-RC[-" + Trim(Str(A_SOCC_ASSIGN_TO_ARRIVE_COL - A_SOCC_ASSIGN_COL)) + "]) * 1440"
                If acReferToServiceCenterTime <> "" Then
                    excelWorksheet.Cells(iRowPosition, A_SOCC_CREATE_TO_REFER_COL) = "=+(RC[+" + Trim(Str(A_SOCC_REFERED_TIME_COL - A_SOCC_CREATE_TO_REFER_COL)) + "]-RC[-" + Trim(Str(A_SOCC_CREATE_TO_REFER_COL - A_FIRST_CALL_COL)) + "]) * 1440"
                End If
                
                acTreeUnit = "=+IF(IsBlank(RC[+" + Trim(Str(A_SC_REPORT_COL - A_RESTORED_BY_COL)) + "]), ""SOCC"", IF(RC[-" + Trim(Str(A_RESTORED_BY_COL - A_ENERGIZED_DATE_COL)) + "] < RC[+" + Trim(Str(A_SC_REPORT_COL - A_RESTORED_BY_COL)) + "], ""SOCC"", ""SC"")) "
                
                '* Check restored by column
                excelWorksheet.Cells(iRowPosition, A_RESTORED_BY_COL) = "=+IF(IsBlank(RC[+" + Trim(Str(A_SC_REPORT_COL - A_RESTORED_BY_COL)) + "]), ""SOCC"", IF(RC[-" + Trim(Str(A_RESTORED_BY_COL - A_ENERGIZED_DATE_COL)) + "] < RC[+" + Trim(Str(A_SC_REPORT_COL - A_RESTORED_BY_COL)) + "], ""SOCC"", ""SC"")) "
                       
                iRowPosition = iRowPosition + 1
                adoIsrList.MoveNext
                
                'Update the progress bar
                i = i + 1
                frmMain.ProgressBar1.Value = i + 1
                frmMain.sbStatusBar.Panels(1).Text = "Processing " + Trim(excelWorksheet.Name) + " for job : " + acIsrNo
            Loop
        
            '* Resize all the columns to fit the data
            'For x = 1 To 3
            '    Set excelWorksheet = excelWorkBook.Worksheets(x)
                excelWorksheet.Select
                excelExcelApp.Cells.Select
                excelExcelApp.Selection.Columns.AutoFit
                excelExcelApp.Range("A2").Select
            'Next
            
        End If
    End If
    frmMain.ProgressBar1.Value = 1
    frmMain.sbStatusBar.Panels(1).Text = "Saving excel file... "
    
    adoIsrList.Close
                
    excelWorkBook.SaveAs acExcelFile, , , , , , xlShared, xlUserResolution
    excelWorkBook.Close False
    'excelExcelApp.Quit
    adoConnection.Close

    frmMain.sbStatusBar.Panels(1).Text = "Done. "
    frmMain.MousePointer = vbDefault
    frmMain.cmdGoButton.Enabled = True
    frmMain.cmdClose.Enabled = True
    
End Sub

Rem ****************************************************************************
Rem *
Rem * Function Name:    CreateWorksheetHeader
Rem * Paramters    :    excelWorksheet - Worksheet object to create the headers on.
Rem *
Rem * Description   : Creates the header columns for the given worksheet.
Rem ****************************************************************************
Private Sub CreateWorksheetHeader(excelWorksheet As Excel.Worksheet, excelWorkBook As Excel.Workbook, excelApplication As Excel.Application, iHeaderType As Integer)

    '* First set up the basic configuration for the spreadsheet
    excelApplication.Cells.Select
    With excelApplication.Selection.Font
        .Name = "Arial"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    excelApplication.Range("A1").Select
    With excelApplication.Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = True
    End With
    excelApplication.Range("A2").Select
    excelApplication.ActiveWindow.FreezePanes = True
    excelApplication.Rows("1:1").Select
    excelApplication.Selection.Font.Bold = True

    '* Set the color different for SOCC
    excelApplication.Range("A1", "W1").Select
    excelApplication.Selection.Interior.ColorIndex = 37
    
    '* Set the color for Service Center
    excelApplication.Range("X1", "AI1").Select
    excelApplication.Selection.Interior.ColorIndex = 40
    
    '* Set the color for Tree
    excelApplication.Range("AJ1", "AX1").Select
    excelApplication.Selection.Interior.ColorIndex = 42
    
        
    excelWorksheet.Cells(1, A_CAD_ID_COL) = "CAD-ID"
    excelWorksheet.Cells(1, A_FINAL_SERVICE_CODE_COL) = "Final Problem Code"
    excelWorksheet.Cells(1, A_JOB_STATUS_ACTIVE) = "Job Status"
    excelWorksheet.Cells(1, A_PRIORITY_FLAG_COL) = "Priority"
    excelWorksheet.Cells(1, A_CIRCUIT_COL) = "Circuit"
    excelWorksheet.Cells(1, A_ZONE_COL) = "Zone"
    
    excelWorksheet.Cells(1, A_REPORT_TOTAL_COL) = "# Field Reports"
    excelWorksheet.Columns(A_REPORT_TOTAL_COL).Select
    excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, A_FIRST_CALL_COL) = "OMS Created"
    excelWorksheet.Columns(A_FIRST_CALL_COL).Select
    excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
    
'    excelWorksheet.Cells(1, A_CAD_ID_CREATED_COL) = "CAD-ID Created"
'    excelWorksheet.Columns(A_CAD_ID_CREATED_COL).Select
'    excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
    
    excelWorksheet.Cells(1, A_ENERGIZED_DATE_COL) = "Service Restored"
    excelWorksheet.Columns(A_ENERGIZED_DATE_COL).Select
    excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
    
    excelWorksheet.Cells(1, A_CREATE_TO_RESTORE_COL) = "Create -> Restore"
    excelWorksheet.Columns(A_CREATE_TO_RESTORE_COL).Select
    excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, A_RESTORED_BY_COL) = "Restored by"
    
    excelWorksheet.Cells(1, A_SOCC_ASSIGN_COL) = "Assign (ST2)"
    excelWorksheet.Columns(A_SOCC_ASSIGN_COL).Select
    excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
    
    excelWorksheet.Cells(1, A_SOCC_FIRST_CREW_COL) = "Initial Crew (ST2)"
    
    excelWorksheet.Cells(1, A_SOCC_ARRIVED_COL) = "Arrvied (ST2)"
    excelWorksheet.Columns(A_SOCC_ARRIVED_COL).Select
    excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
    
    excelWorksheet.Cells(1, A_SOCC_REPORTING_COL) = "Reporting (ST2)"
    excelWorksheet.Columns(A_SOCC_REPORTING_COL).Select
    excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
    
    'excelWorksheet.Cells(1, A_CREATE_TO_CAD_COL) = "Create -> CAD-ID (ST2)"
    'excelWorksheet.Columns(A_CREATE_TO_CAD_COL).Select
    'excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, A_SOCC_CREATE_TO_ARRIVE_COL) = "Create -> Arrived (ST2)"
    excelWorksheet.Columns(A_SOCC_CREATE_TO_ARRIVE_COL).Select
    excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, A_SOCC_CREATE_TO_REPORT_COL) = "Create -> Report (ST2)"
    excelWorksheet.Columns(A_SOCC_CREATE_TO_REPORT_COL).Select
    excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, A_SOCC_ARRIVED_TO_REPORT_COL) = "Arrived -> Report (ST2)"
    excelWorksheet.Columns(A_SOCC_ARRIVED_TO_REPORT_COL).Select
    excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, A_SOCC_ASSIGN_TO_ARRIVE_COL) = "Assign -> Arrived"
    excelWorksheet.Columns(A_SOCC_ASSIGN_TO_ARRIVE_COL).Select
    excelApplication.Selection.NumberFormat = "0"

    excelWorksheet.Cells(1, A_SOCC_REFERED_FLAG_COL) = "Referred from ST2?"
    excelWorksheet.Cells(1, A_SOCC_REFERED_TO_COL) = "Referred where?"
    
    excelWorksheet.Cells(1, A_SOCC_CREATE_TO_REFER_COL) = "Create -> Referred"
    excelWorksheet.Columns(A_SOCC_CREATE_TO_REFER_COL).Select
    excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, A_SOCC_REFERED_TIME_COL) = "Referred to?"
    excelWorksheet.Columns(A_SOCC_REFERED_TIME_COL).Select
    excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
    
    excelWorksheet.Cells(1, A_SC_PRIORITY_REFER_TIME_COL) = "Time Referred to SC"
    excelWorksheet.Columns(A_SC_PRIORITY_REFER_TIME_COL).Select
    excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
    
    excelWorksheet.Cells(1, A_SC_REFER_FROM_COL) = "Referred from (SC)"
    excelWorksheet.Cells(1, A_SC_REFER_TO_COL) = "Referred to (SC)"
    
    excelWorksheet.Cells(1, A_SC_UNIT_COL) = "Referred Crew (SC)"
    
    excelWorksheet.Cells(1, A_SC_ASSIGN_COL) = "Assigned (SC)"
    excelWorksheet.Columns(A_SC_ASSIGN_COL).Select
    excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
    
    excelWorksheet.Cells(1, A_SC_ARRIVE_COL) = "Arrived (SC)"
    excelWorksheet.Columns(A_SC_ARRIVE_COL).Select
    excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
    
    excelWorksheet.Cells(1, A_SC_REPORT_COL) = "Reporting (SC)"
    excelWorksheet.Columns(A_SC_REPORT_COL).Select
    excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
    
    excelWorksheet.Cells(1, A_SC_CREATE_TO_ARRIVE_COL) = "Create -> Arrived (SC)"
    excelWorksheet.Columns(A_SC_CREATE_TO_ARRIVE_COL).Select
    excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, A_SC_CREATE_TO_REPORT_COL) = "Create -> Report (SC)"
    excelWorksheet.Columns(A_SC_CREATE_TO_REPORT_COL).Select
    excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, A_SC_REFER_TO_ARRIVE_COL) = "Refer -> Arrive (SC)"
    excelWorksheet.Columns(A_SC_REFER_TO_ARRIVE_COL).Select
    excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, A_SC_REFER_TO_REPORT_COL) = "Refer -> Report (SC)"
    excelWorksheet.Columns(A_SC_REFER_TO_REPORT_COL).Select
    excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, A_SC_ARRIVE_TO_REPORT_COL) = "Arrive -> Report (SC)"
    excelWorksheet.Columns(A_SC_ARRIVE_TO_REPORT_COL).Select
    excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, A_TREE_REFER_TIME_COL) = "Time Referred to Tree screen (Tree)"
    excelWorksheet.Columns(A_TREE_REFER_TIME_COL).Select
    excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
    
    excelWorksheet.Cells(1, A_TREE_REFER_TO_COL) = "Refer to (Tree)"
    excelWorksheet.Cells(1, A_TREE_UNIT_COL) = "Referred Crew (Tree)"
    
    excelWorksheet.Cells(1, A_TREE_ASSIGN_COL) = "Assigned (Tree)"
    excelWorksheet.Columns(A_TREE_ASSIGN_COL).Select
    excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
    
    excelWorksheet.Cells(1, A_TREE_ARRIVE_COL) = "Arrived (Tree)"
    excelWorksheet.Columns(A_TREE_ARRIVE_COL).Select
    excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
    
    excelWorksheet.Cells(1, A_TREE_REPORT_COL) = "Reporting (Tree)"
    excelWorksheet.Columns(A_TREE_REPORT_COL).Select
    excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
    
    excelWorksheet.Cells(1, A_TREE_CREATE_TO_ARRIVE_COL) = "Create -> Arrived (Tree)"
    excelWorksheet.Columns(A_TREE_CREATE_TO_ARRIVE_COL).Select
    excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, A_TREE_CREATE_TO_REPORT_COL) = "Create -> Report (Tree)"
    excelWorksheet.Columns(A_TREE_CREATE_TO_REPORT_COL).Select
    excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, A_TREE_REFER_TO_ARRIVE_COL) = "Refer -> Arrive(Tree)"
    excelWorksheet.Columns(A_TREE_REFER_TO_ARRIVE_COL).Select
    excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, A_TREE_REFER_TO_REPORT_COL) = "Refer -> Report (Tree)"
    excelWorksheet.Columns(A_TREE_REFER_TO_REPORT_COL).Select
    excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, A_TREE_ARRIVE_TO_REPORT_COL) = "Arrive -> Report (Tree)"
    excelWorksheet.Columns(A_TREE_ARRIVE_TO_REPORT_COL).Select
    excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, A_JOB_DISPOSE_TIME_COL) = "Disposed Time"
    excelWorksheet.Columns(A_JOB_DISPOSE_TIME_COL).Select
    excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
    
    
    excelWorksheet.Cells(1, A_TOTAL_CALLS_COL) = "Total Calls"
    excelWorksheet.Columns(A_TOTAL_CALLS_COL).Select
    excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, A_TOTAL_CMO_COL) = "Total Customers Minutes Out"
    excelWorksheet.Columns(A_TOTAL_CMO_COL).Select
    excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, A_TOTAL_CUSTOMERS_COL) = "Total Customers"
    excelWorksheet.Columns(A_TOTAL_CUSTOMERS_COL).Select
    excelApplication.Selection.NumberFormat = "0"

    
End Sub






