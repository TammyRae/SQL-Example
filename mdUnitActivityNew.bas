Attribute VB_Name = "mdUnitActivityNew"
Option Explicit

'******************************************
'*** Excel column constants for report ****
'******************************************
Public Const PRIORITY_1_LIST As String = "FBLD, LO, LPAR, META, PBRO, PDOW, PFIR, PHBC, POLB, TDWP, TXFG, TXFP, TXHB, UCCC, UCCP, UCEX, UTUW, WAAH, WAAP, WAIG, WAIT, WBAH, WBAP, WBAT, WDLP"

Public Const A_CAD_ID_COLUMN As Integer = 1
Public Const A_UNIT_COLUMN As Integer = 2
Public Const A_PROBLEM_CODE_COLUMN As Integer = 3
Public Const A_REFER_TO_COLUMN As Integer = 4
Public Const A_CREATED_DATE_COLUMN As Integer = 5
Public Const A_DISPATCH_DATE_COLUMN As Integer = 6
Public Const A_ACCEPTED_DATE_COLUMN As Integer = 7
Public Const A_ENROUTE_DATE_COLUMN As Integer = 8
Public Const A_ARRIVED_DATE_COLUMN As Integer = 9
Public Const A_REPORT_DATE_COLUMN As Integer = 10
Public Const A_CLEARED_DATE_COLUMN As Integer = 11
Public Const A_CREATE_TO_ASSIGN_COLUMN As Integer = 12
Public Const A_ASSIGNED_TO_ARRIVED_COLUMN As Integer = 13
Public Const A_CREATE_TO_CLEARED_COLUMN As Integer = 14
Public Const A_CREATE_TO_ARRIVED_COLUMN As Integer = 15
Public Const A_ENROUTE_TO_ARRIVED_COLUMN As Integer = 16
Public Const A_ENROUTE_TO_CLEARED_COLUMN As Integer = 17
Public Const A_ARRIVED_TO_CLEARED_COLUMN As Integer = 18
Public Const A_TOTAL_CALLS_COLUMN As Integer = 19
Public Const A_TOTAL_CUSTOMERS_COLUMN As Integer = 20

Public Const B_CAD_ID_COLUMN As Integer = 1
Public Const B_PROBLEM_CODE_COLUMN As Integer = 2
Public Const B_INITIAL_UNIT_COLUMN As Integer = 3
Public Const B_CREATED_DATE_COLUMN As Integer = 4
Public Const B_INITIAL_ENROUTE_DATE_COLUMN As Integer = 5
Public Const B_INITIAL_ARRIVED_DATE_COLUMN As Integer = 6
Public Const B_INITIAL_CLEARED_DATE_COLUMN As Integer = 7
Public Const B_RESTORED_BY_SOCC As Integer = 8
Public Const B_INITIAL_CREATE_TO_CLEARED_COLUMN As Integer = 9
Public Const B_INITIAL_CREATE_TO_ARRIVED_COLUMN As Integer = 10
Public Const B_INITIAL_ARRIVED_TO_CLEARED_COLUMN As Integer = 11
Public Const B_INITIAL_ENROUTE_TO_ARRIVED_COLUMN As Integer = 12
Public Const B_RESTORE_DATE_COLUMN As Integer = 13
Public Const B_REFER_TO_COLUMN As Integer = 14
Public Const B_FINAL_UNIT_COLUMN As Integer = 15
Public Const B_FINAL_RECEIPT_DATE_COLUMN As Integer = 16
Public Const B_FINAL_ENROUTE_DATE_COLUMN As Integer = 17
Public Const B_FINAL_ARRIVE_DATE_COLUMN As Integer = 18
Public Const B_FINAL_CLEARED_DATE_COLUMN As Integer = 19
Public Const B_FINAL_CREATE_TO_CLEARED_COLUMN As Integer = 20
Public Const B_FINAL_RECEIPT_TO_ARRIVE_COLUMN As Integer = 21
Public Const B_FINAL_ARRIVE_TO_CLEARED_COLUMN As Integer = 22
Public Const B_FINAL_ENROUTE_TO_ARRIVED_COLUMN As Integer = 23
Public Const B_FINAL_CREATE_TO_RESTORE_COLUMN As Integer = 24
Public Const B_TOTAL_CALLS_COLUMN As Integer = 25
Public Const B_TOTAL_CUSTOMERS_COLUMN As Integer = 26


Rem ****************************************************************************
Rem *
Rem * Function Name:    GenerateNewUnitActivityReport
Rem * Paramters    :    acUserName      - User name to login to cad database
Rem *                   acPassword      - Password to use to login to database.
Rem *                   acDBInstance    - The CAD db instance
Rem *                   acOutputName    - File name to same output (spreadsheet) to.
Rem *
Rem * Description   : Generate the unit activity report
Rem ****************************************************************************
Sub GenerateNewUnitActivityReport(acUserName As String, acPassword As String, _
                                acDBInstance As String, acExcelFile As String, acReportType As String)
    Dim acSQLText As String, acIsrNo As String, acUnitNo As String
    Dim acCellData As String, acPreviousUnitNo As String
    Dim acDispatchTime As String, acReportingTime As String
    Dim acAcceptedTime As String, acEnrouteTime As String
    Dim acArrivedTime As String, acClearedTime As String
    Dim acCreatedTime As String, blPriority1 As Boolean
    Dim acProblemCode As String, iPriorityPosition As Integer, iPriority2Position As Integer
    Dim i As Integer, x As Integer, iTotalRows As Double, iRowPosition As Double
    Dim iCrewTotal As Integer, iTimeDifference As Integer
    Dim iCounter As Integer, acReferUnit As String
    Dim excelExcelApp As Excel.Application ' This is the excel program
    Dim excelWorkBook As Excel.Workbook ' This is the work book
    Dim excelWorksheet As Excel.Worksheet ' This is the sheet
    Dim adoConnection As ADODB.Connection
    Dim adoUnitList As ADODB.Recordset
    Dim adoStatusTime As ADODB.Recordset
    Dim adoReferUnit As ADODB.Recordset
    
    Set adoConnection = New ADODB.Connection
    Set adoUnitList = New ADODB.Recordset
    Set adoStatusTime = New ADODB.Recordset
    Set adoReferUnit = New ADODB.Recordset
    
    'connect to the Oracle database
    frmMain.MousePointer = vbHourglass
    frmMain.sbStatusBar.Panels(1).Text = "Opening connection to database..."
    frmMain.cmdGoButton.Enabled = False
    frmMain.cmdClose.Enabled = False
    adoConnection.Open "Provider=MSDASQL.1;Persist Security Info=False;User ID=" & acUserName & ";pwd=" & acPassword & ";Data Source=" & acDBInstance & ";"
    adoConnection.CursorLocation = adUseClient
    adoConnection.CommandTimeout = 60

    ' If no directory is specified, put in current directory
    If (InStr(acExcelFile, "\") = 0) Then
        acExcelFile = CurDir() + "\" + Trim(acExcelFile)
    End If
    
    StartExcel excelExcelApp
    Set excelWorkBook = excelExcelApp.Workbooks.Add
    
    acSQLText = " select ISR.ISR_NO, UNIT, ISR.PRIORITY, P_FILTER1, "
    acSQLText = acSQLText + " DECODE(ISR.init_service_code, 'MCC OFF', 'EROU', 'MCC', 'EROU', 'M-DISREM', 'EROU', 'PL-DIS', 'EROU', 'POLE-REC', 'EROU', 'POLE-SAF', 'EROU', 'TDL-ONLY', 'EROU', 'TDLPM', 'EROU', 'TLAT', 'EROU', 'TDPM', 'EROU', 'DPNP', 'EROU', 'DPSF', 'EROU', 'RCPP', 'EROU', 'DPMI', 'EROU', (SELECT CALL_TYPE_CD FROM CLUE@fmsrpt_ompr11_l WHERE CLUE_CD = ISR.INIT_SERVICE_CODE)) CALL_TYPE_CD,  "
    acSQLText = acSQLText + " ISR.init_service_code, D11_UNIT_ACTIVITY.AGENCY_CODE, "
    acSQLText = acSQLText + " decode( isr.restoration_date, ' ', '', "
    acSQLText = acSQLText + " to_char( to_date(isr.restoration_date || isr.restoration_time, 'YYYYMMDDHH24MISS'), 'MM/DD/YYYY HH24:MI:SS')) RESTORE_DATE, "
    acSQLText = acSQLText + " TO_CHAR( TO_DATE(D11_START_DATE || D11_START_TIME, 'YYYYMMDDHH24MISS'), 'MM/DD/YYYY HH24:MI:SS') FORMATED_START_TIME, "
    acSQLText = acSQLText + " TO_CHAR( TO_DATE(ISR.CREATED_DATE || ISR.CREATED_TIME, 'YYYYMMDDHH24MISS'), 'MM/DD/YYYY HH24:MI:SS') CREATION_TIME, ISR.UDF11 TOTAL_CALLS, ISR.UDF7 CUSTOMERS_AFFECTED "
    acSQLText = acSQLText + " From d11_unit_activity, ISR "
    acSQLText = acSQLText + " Where ISR.ISR_NO = d11_unit_activity.ISR_NO "
    acSQLText = acSQLText + " and b07_status_code = 'RP' "
    acSQLText = acSQLText + " and d11_start_date >= '" + Format(frmMain.lstDatePickerFrom.Value, "YYYYMMDD") + "' "
    acSQLText = acSQLText + " and d11_end_date <= '" + Format(frmMain.lstDatePickerTo.Value, "YYYYMMDD") + "' "
    acSQLText = acSQLText + " and d11_unit_activity.agency_Code = 'ELEC' "
    acSQLText = acSQLText + " and ((d11_unit_activity.unit > '199' and d11_unit_activity.unit < '251') or d11_unit_activity.unit = '637')"
    acSQLText = acSQLText + " order by CALL_TYPE_CD, UNIT, D11_START_DATE, D11_START_TIME "

    frmMain.sbStatusBar.Panels(1).Text = "Running query..."
    adoUnitList.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText

    If Not (adoUnitList.BOF And adoUnitList.EOF) Then
        iTotalRows = adoUnitList.RecordCount
          
        If iTotalRows > 0 Then
            For iCounter = excelWorkBook.Worksheets.Count To 3  '** There is usually 3 worksheets open by default in excel
                excelWorkBook.Worksheets.Add
            Next
            
            ' Create each worksheet and the column headers.
            Set excelWorksheet = excelWorkBook.Worksheets(1)
            excelWorksheet.Name = "Worksheet Restoration"
            excelWorksheet.Select
            CreateWorksheetHeader excelWorksheet, excelWorkBook, excelExcelApp, 2
            Set excelWorksheet = excelWorkBook.Worksheets(2)
            excelWorksheet.Name = "Worksheet Priority - Rest"
            excelWorksheet.Select
            CreateWorksheetHeader excelWorksheet, excelWorkBook, excelExcelApp, 1
            Set excelWorksheet = excelWorkBook.Worksheets(3)
            excelWorksheet.Name = "Routine All"
            excelWorksheet.Select
            CreateWorksheetHeader excelWorksheet, excelWorkBook, excelExcelApp, 1
            Set excelWorksheet = excelWorkBook.Worksheets(4)
            excelWorksheet.Select
            excelWorksheet.Name = "Street Lights"
            CreateWorksheetHeader excelWorksheet, excelWorkBook, excelExcelApp, 1
           
            'Update the progress bar
            frmMain.ProgressBar1.Max = iTotalRows + 1
            frmMain.ProgressBar1.Min = 1
            frmMain.ProgressBar1.Visible = True
            
            i = 0
            iPriorityPosition = 2
            iPriority2Position = 2
            iRowPosition = 2
            adoUnitList.MoveFirst
            acPreviousUnitNo = adoUnitList.Fields("UNIT")
            iCrewTotal = 0
            
            Set excelWorksheet = excelWorkBook.Worksheets(1)
            Do While Not adoUnitList.EOF
                acProblemCode = adoUnitList.Fields.Item("INIT_SERVICE_CODE")
                blPriority1 = (InStr(PRIORITY_1_LIST, acProblemCode) > 0)
                If adoUnitList.Fields("CALL_TYPE_CD") = "ELEC" Then
                    If excelWorksheet.Name <> "Worksheet Restoration" And blPriority1 Then
                        Set excelWorksheet = excelWorkBook.Worksheets(1)
                        acPreviousUnitNo = adoUnitList.Fields("UNIT")
                        iRowPosition = iPriorityPosition
                    ElseIf excelWorksheet.Name <> "Worksheet Priority - Rest" And Not blPriority1 Then
                        Set excelWorksheet = excelWorkBook.Worksheets(2)
                        acPreviousUnitNo = adoUnitList.Fields("UNIT")
                        iCrewTotal = 0
                        iRowPosition = iPriority2Position
                    End If
                ElseIf excelWorksheet.Name <> "Routine All" And adoUnitList.Fields("CALL_TYPE_CD") = "EROU" Then
                    Set excelWorksheet = excelWorkBook.Worksheets(3)
                    acPreviousUnitNo = adoUnitList.Fields("UNIT")
                    iCrewTotal = 0
                    iRowPosition = 2
                ElseIf excelWorksheet.Name <> "Street Lights" And adoUnitList.Fields("CALL_TYPE_CD") = "STRL" Then
                    Set excelWorksheet = excelWorkBook.Worksheets(4)
                    acPreviousUnitNo = adoUnitList.Fields("UNIT")
                    iCrewTotal = 0
                    iRowPosition = 2
                End If

                acIsrNo = adoUnitList.Fields.Item("ISR_NO")
                acReportingTime = adoUnitList.Fields("FORMATED_START_TIME")
                acCreatedTime = adoUnitList.Fields("CREATION_TIME")
                acUnitNo = adoUnitList.Fields("UNIT")
                
                If blPriority1 Then
                    excelWorksheet.Cells(iRowPosition, B_CAD_ID_COLUMN) = acIsrNo
                    excelWorksheet.Cells(iRowPosition, B_PROBLEM_CODE_COLUMN) = adoUnitList.Fields("INIT_SERVICE_CODE")
                    excelWorksheet.Cells(iRowPosition, B_INITIAL_UNIT_COLUMN) = acUnitNo
                    excelWorksheet.Cells(iRowPosition, B_CREATED_DATE_COLUMN) = acCreatedTime
                    excelWorksheet.Cells(iRowPosition, B_RESTORE_DATE_COLUMN) = adoUnitList.Fields("RESTORE_DATE")
                    excelWorksheet.Cells(iRowPosition, B_REFER_TO_COLUMN) = adoUnitList.Fields("P_FILTER1")
                    excelWorksheet.Cells(iRowPosition, B_TOTAL_CALLS_COLUMN) = adoUnitList.Fields("TOTAL_CALLS")
                    excelWorksheet.Cells(iRowPosition, B_TOTAL_CUSTOMERS_COLUMN) = adoUnitList.Fields("CUSTOMERS_AFFECTED")
                    excelWorksheet.Cells(iRowPosition, B_RESTORED_BY_SOCC) = "=IF((RC[-" + Trim(Str(B_RESTORED_BY_SOCC - B_INITIAL_CLEARED_DATE_COLUMN)) + "])>=(RC[+" + Trim(Str(B_RESTORE_DATE_COLUMN - B_RESTORED_BY_SOCC)) + "]), ""No"", ""Yes"")"
                    
                    '* Get the referal time
                    If (Len(Trim(adoUnitList.Fields("P_FILTER1"))) > 0) Then
                        acSQLText = "SELECT MAX(TO_DATE((MOD_DATE || MOD_TIME), 'YYYYMMDDHH24MISS')) STATUS_TIME "
                        acSQLText = acSQLText + "FROM ISR_LOG WHERE ISR_NO = '" + acIsrNo + "' "
                        acSQLText = acSQLText + "AND FIELDNAME = 'p_filter1' and FIELDVALUE = '" + adoUnitList.Fields("P_FILTER1") + "' "
                        adoStatusTime.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                        If Not (adoStatusTime.BOF And adoStatusTime.EOF) Then
                            If Not IsNull(adoStatusTime.Fields("STATUS_TIME")) Then
                                acDispatchTime = adoStatusTime.Fields("STATUS_TIME")
                                excelWorksheet.Cells(iRowPosition, B_FINAL_RECEIPT_DATE_COLUMN) = acDispatchTime
                            End If
                        End If
                        adoStatusTime.Close
                    End If
                Else
                    excelWorksheet.Cells(iRowPosition, A_CAD_ID_COLUMN) = acIsrNo
                    excelWorksheet.Cells(iRowPosition, A_PROBLEM_CODE_COLUMN) = adoUnitList.Fields("INIT_SERVICE_CODE")
                    excelWorksheet.Cells(iRowPosition, A_UNIT_COLUMN) = acUnitNo
                    excelWorksheet.Cells(iRowPosition, A_CREATED_DATE_COLUMN) = acCreatedTime
                    excelWorksheet.Cells(iRowPosition, A_REPORT_DATE_COLUMN) = acReportingTime
                    excelWorksheet.Cells(iRowPosition, A_REFER_TO_COLUMN) = adoUnitList.Fields("P_FILTER1")
                    excelWorksheet.Cells(iRowPosition, A_TOTAL_CALLS_COLUMN) = adoUnitList.Fields("TOTAL_CALLS")
                    excelWorksheet.Cells(iRowPosition, A_TOTAL_CUSTOMERS_COLUMN) = adoUnitList.Fields("CUSTOMERS_AFFECTED")
                End If
                     
                If Not blPriority1 Then
                    '* Get the dispatch time
                    acSQLText = "select to_char(max(to_date(d11_start_date || d11_start_time, 'YYYYMMDDHH24MISS')), 'MM/DD/YYYY HH24:MI:SS') STATUS_TIME "
                    acSQLText = acSQLText + " from d11_unit_activity "
                    acSQLText = acSQLText + " where isr_no = '" + acIsrNo + "' "
                    acSQLText = acSQLText + " and b07_status_code = 'DP' and unit = '" + acUnitNo + "' "
                    acSQLText = acSQLText + " and (d11_start_date || d11_start_time) <= '" + Mid(acReportingTime, 7, 4) + Mid(acReportingTime, 1, 2) + Mid(acReportingTime, 4, 2) + Mid(acReportingTime, 12, 2) + Mid(acReportingTime, 15, 2) + Mid(acReportingTime, 18, 2) + "' "
                     
                    adoStatusTime.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                    If Not (adoStatusTime.BOF And adoStatusTime.EOF) Then
                        If Not IsNull(adoStatusTime.Fields("STATUS_TIME")) Then
                            acDispatchTime = adoStatusTime.Fields("STATUS_TIME")
                            excelWorksheet.Cells(iRowPosition, A_DISPATCH_DATE_COLUMN) = acDispatchTime
                        End If
                    End If
                    adoStatusTime.Close
                
                    '* Get the accepted time
                    acSQLText = "select to_char(max(to_date(d11_start_date || d11_start_time, 'YYYYMMDDHH24MISS')), 'MM/DD/YYYY HH24:MI') STATUS_TIME "
                    acSQLText = acSQLText + " from d11_unit_activity "
                    acSQLText = acSQLText + " where isr_no = '" + acIsrNo + "' "
                    acSQLText = acSQLText + " and b07_status_code = 'AC' and unit = '" + acUnitNo + "' "
                    acSQLText = acSQLText + " and (d11_start_date || d11_start_time) <= '" + Mid(acReportingTime, 7, 4) + Mid(acReportingTime, 1, 2) + Mid(acReportingTime, 4, 2) + Mid(acReportingTime, 12, 2) + Mid(acReportingTime, 15, 2) + Mid(acReportingTime, 18, 2) + "' "
                     
                    adoStatusTime.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                    If Not (adoStatusTime.BOF And adoStatusTime.EOF) Then
                        If Not IsNull(adoStatusTime.Fields("STATUS_TIME")) Then
                            acAcceptedTime = adoStatusTime.Fields("STATUS_TIME")
                            excelWorksheet.Cells(iRowPosition, A_ACCEPTED_DATE_COLUMN) = acAcceptedTime
                        End If
                    End If
                    adoStatusTime.Close
                End If
                
                '* Get the enroute time
                acSQLText = "select to_char(max(to_date(d11_start_date || d11_start_time, 'YYYYMMDDHH24MISS')), 'MM/DD/YYYY HH24:MI') STATUS_TIME "
                acSQLText = acSQLText + " from d11_unit_activity "
                acSQLText = acSQLText + " where isr_no = '" + acIsrNo + "' "
                acSQLText = acSQLText + " and b07_status_code = 'ER' and unit = '" + acUnitNo + "' "
                acSQLText = acSQLText + " and (d11_start_date || d11_start_time) <= '" + Mid(acReportingTime, 7, 4) + Mid(acReportingTime, 1, 2) + Mid(acReportingTime, 4, 2) + Mid(acReportingTime, 12, 2) + Mid(acReportingTime, 15, 2) + Mid(acReportingTime, 18, 2) + "' "
                     
                adoStatusTime.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                If Not (adoStatusTime.BOF And adoStatusTime.EOF) Then
                    If Not IsNull(adoStatusTime.Fields("STATUS_TIME")) Then
                        acEnrouteTime = adoStatusTime.Fields("STATUS_TIME")
                        If blPriority1 Then
                            excelWorksheet.Cells(iRowPosition, B_INITIAL_ENROUTE_DATE_COLUMN) = acEnrouteTime
                        Else
                            excelWorksheet.Cells(iRowPosition, A_ENROUTE_DATE_COLUMN) = acEnrouteTime
                        End If
                    End If
                End If
                adoStatusTime.Close
                
                '* Get the arrived time
                acSQLText = "select to_char(max(to_date(d11_start_date || d11_start_time, 'YYYYMMDDHH24MISS')), 'MM/DD/YYYY HH24:MI') STATUS_TIME "
                acSQLText = acSQLText + " from d11_unit_activity "
                acSQLText = acSQLText + " where isr_no = '" + acIsrNo + "' "
                acSQLText = acSQLText + " and b07_status_code = 'AR' and unit = '" + acUnitNo + "' "
                acSQLText = acSQLText + " and (d11_start_date || d11_start_time) <= '" + Mid(acReportingTime, 7, 4) + Mid(acReportingTime, 1, 2) + Mid(acReportingTime, 4, 2) + Mid(acReportingTime, 12, 2) + Mid(acReportingTime, 15, 2) + Mid(acReportingTime, 18, 2) + "' "
                     
                adoStatusTime.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                If Not (adoStatusTime.BOF And adoStatusTime.EOF) Then
                    If Not IsNull(adoStatusTime.Fields("STATUS_TIME")) Then
                        acArrivedTime = adoStatusTime.Fields("STATUS_TIME")
                        If blPriority1 Then
                            excelWorksheet.Cells(iRowPosition, B_INITIAL_ARRIVED_DATE_COLUMN) = acArrivedTime
                        Else
                            excelWorksheet.Cells(iRowPosition, A_ARRIVED_DATE_COLUMN) = acArrivedTime
                        End If
                    End If
                End If
                adoStatusTime.Close
                
                '* Get the cleared time
                acSQLText = "select to_char(min(TO_DATE(c07_start_date || c07_start_time, 'YYYYMMDDHH24MISS')), 'MM/DD/YYYY HH24:MI') STATUS_TIME "
                acSQLText = acSQLText + " from c07_isr_activity "
                acSQLText = acSQLText + " where isr_no = '" + acIsrNo + "' "
                acSQLText = acSQLText + " and c07_activity_code = 'CL' "
                acSQLText = acSQLText + " and (c07_start_date || c07_start_time) >= '" + Mid(acReportingTime, 7, 4) + Mid(acReportingTime, 1, 2) + Mid(acReportingTime, 4, 2) + Mid(acReportingTime, 12, 2) + Mid(acReportingTime, 15, 2) + Mid(acReportingTime, 18, 2) + "' "
               
                adoStatusTime.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                If Not (adoStatusTime.BOF And adoStatusTime.EOF) Then
                    If Not IsNull(adoStatusTime.Fields("STATUS_TIME")) Then
                        acClearedTime = adoStatusTime.Fields("STATUS_TIME")
                        If blPriority1 Then
                            excelWorksheet.Cells(iRowPosition, B_INITIAL_CLEARED_DATE_COLUMN) = acClearedTime
                        Else
                            excelWorksheet.Cells(iRowPosition, A_CLEARED_DATE_COLUMN) = acClearedTime
                        End If
                    End If
                End If
                adoStatusTime.Close
                
                '* Put the formulas for the time differences columns
                If blPriority1 Then
                    excelWorksheet.Cells(iRowPosition, B_INITIAL_CREATE_TO_CLEARED_COLUMN) = "=+(RC[-" + Trim(Str(B_INITIAL_CREATE_TO_CLEARED_COLUMN - B_INITIAL_CLEARED_DATE_COLUMN)) + "]-RC[-" + Trim(Str(B_INITIAL_CREATE_TO_CLEARED_COLUMN - B_CREATED_DATE_COLUMN)) + "])*1440"
                    excelWorksheet.Cells(iRowPosition, B_INITIAL_CREATE_TO_ARRIVED_COLUMN) = "=+(RC[-" + Trim(Str(B_INITIAL_CREATE_TO_ARRIVED_COLUMN - B_INITIAL_ARRIVED_DATE_COLUMN)) + "]-RC[-" + Trim(Str(B_INITIAL_CREATE_TO_ARRIVED_COLUMN - B_CREATED_DATE_COLUMN)) + "])*1440"
                    excelWorksheet.Cells(iRowPosition, B_INITIAL_ARRIVED_TO_CLEARED_COLUMN) = "=+(RC[-" + Trim(Str(B_INITIAL_ARRIVED_TO_CLEARED_COLUMN - B_INITIAL_CLEARED_DATE_COLUMN)) + "]-RC[-" + Trim(Str(B_INITIAL_ARRIVED_TO_CLEARED_COLUMN - B_INITIAL_ARRIVED_DATE_COLUMN)) + "])*1440"
                    excelWorksheet.Cells(iRowPosition, B_INITIAL_ENROUTE_TO_ARRIVED_COLUMN) = "=+(RC[-" + Trim(Str(B_INITIAL_ENROUTE_TO_ARRIVED_COLUMN - B_INITIAL_ARRIVED_DATE_COLUMN)) + "]-RC[-" + Trim(Str(B_INITIAL_ENROUTE_TO_ARRIVED_COLUMN - B_INITIAL_ENROUTE_DATE_COLUMN)) + "])*1440"
                    excelWorksheet.Cells(iRowPosition, B_FINAL_CREATE_TO_RESTORE_COLUMN) = "=+(RC[-" + Trim(Str(B_FINAL_CREATE_TO_RESTORE_COLUMN - B_RESTORE_DATE_COLUMN)) + "]-RC[-" + Trim(Str(B_FINAL_CREATE_TO_RESTORE_COLUMN - B_CREATED_DATE_COLUMN)) + "])*1440"
                        
                    '* Get the referal crew
                    acSQLText = "select UNIT, "
                    acSQLText = acSQLText + "TO_CHAR( TO_DATE(D11_START_DATE || D11_START_TIME, 'YYYYMMDDHH24MISS'), 'MM/DD/YYYY HH24:MI:SS') FORMATED_START_TIME "
                    acSQLText = acSQLText + "from d11_unit_activity d1 "
                    acSQLText = acSQLText + "where isr_no = '" + acIsrNo + "' "
                    acSQLText = acSQLText + "and (d1.unit < '200' or d1.unit > '251') "
                    acSQLText = acSQLText + "and b07_status_code = 'RP' "
                    acSQLText = acSQLText + "and (d11_start_date || d11_start_time) = "
                    acSQLText = acSQLText + "(select max(d11_start_date || d11_start_time) from d11_unit_activity d2 "
                    acSQLText = acSQLText + "where d1.isr_no = d2.isr_no and "
                    acSQLText = acSQLText + "(d2.unit < '200' or d2.unit > '251') and b07_status_code = 'RP') "
                    
                    adoReferUnit.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                    If Not adoReferUnit.EOF Then
                        acReferUnit = adoReferUnit.Fields.Item("UNIT")
                        acReportingTime = adoReferUnit.Fields("FORMATED_START_TIME")
                        excelWorksheet.Cells(iRowPosition, B_FINAL_UNIT_COLUMN) = acReferUnit
                        
                        '* Get the enroute time
                        acSQLText = "select to_char(max(to_date(d11_start_date || d11_start_time, 'YYYYMMDDHH24MISS')), 'MM/DD/YYYY HH24:MI') STATUS_TIME "
                        acSQLText = acSQLText + " from d11_unit_activity "
                        acSQLText = acSQLText + " where isr_no = '" + acIsrNo + "' "
                        acSQLText = acSQLText + " and b07_status_code = 'ER' and unit = '" + acReferUnit + "' "
                        acSQLText = acSQLText + " and (d11_start_date || d11_start_time) <= '" + Mid(acReportingTime, 7, 4) + Mid(acReportingTime, 1, 2) + Mid(acReportingTime, 4, 2) + Mid(acReportingTime, 12, 2) + Mid(acReportingTime, 15, 2) + Mid(acReportingTime, 18, 2) + "' "
                             
                        adoStatusTime.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                        If Not (adoStatusTime.BOF And adoStatusTime.EOF) Then
                            If Not IsNull(adoStatusTime.Fields("STATUS_TIME")) Then
                                acEnrouteTime = adoStatusTime.Fields("STATUS_TIME")
                                excelWorksheet.Cells(iRowPosition, B_FINAL_ENROUTE_DATE_COLUMN) = acEnrouteTime
                            End If
                        End If
                        adoStatusTime.Close
                        
                        '* Get the arrived time
                        acSQLText = "select to_char(max(to_date(d11_start_date || d11_start_time, 'YYYYMMDDHH24MISS')), 'MM/DD/YYYY HH24:MI') STATUS_TIME "
                        acSQLText = acSQLText + " from d11_unit_activity "
                        acSQLText = acSQLText + " where isr_no = '" + acIsrNo + "' "
                        acSQLText = acSQLText + " and b07_status_code = 'AR' and unit = '" + acReferUnit + "' "
                        acSQLText = acSQLText + " and (d11_start_date || d11_start_time) <= '" + Mid(acReportingTime, 7, 4) + Mid(acReportingTime, 1, 2) + Mid(acReportingTime, 4, 2) + Mid(acReportingTime, 12, 2) + Mid(acReportingTime, 15, 2) + Mid(acReportingTime, 18, 2) + "' "
                             
                        adoStatusTime.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                        If Not (adoStatusTime.BOF And adoStatusTime.EOF) Then
                            If Not IsNull(adoStatusTime.Fields("STATUS_TIME")) Then
                                acArrivedTime = adoStatusTime.Fields("STATUS_TIME")
                                excelWorksheet.Cells(iRowPosition, B_FINAL_ARRIVE_DATE_COLUMN) = acArrivedTime
                            End If
                        End If
                        adoStatusTime.Close
                        
                        '* Get the cleared time
                        acSQLText = "select to_char(min(TO_DATE(c07_start_date || c07_start_time, 'YYYYMMDDHH24MISS')), 'MM/DD/YYYY HH24:MI') STATUS_TIME "
                        acSQLText = acSQLText + " from c07_isr_activity "
                        acSQLText = acSQLText + " where isr_no = '" + acIsrNo + "' "
                        acSQLText = acSQLText + " and c07_activity_code = 'CL' "
                        acSQLText = acSQLText + " and (c07_start_date || c07_start_time) >= '" + Mid(acReportingTime, 7, 4) + Mid(acReportingTime, 1, 2) + Mid(acReportingTime, 4, 2) + Mid(acReportingTime, 12, 2) + Mid(acReportingTime, 15, 2) + Mid(acReportingTime, 18, 2) + "' "
                        
                        adoStatusTime.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                        If Not (adoStatusTime.BOF And adoStatusTime.EOF) Then
                            If Not IsNull(adoStatusTime.Fields("STATUS_TIME")) Then
                                acClearedTime = adoStatusTime.Fields("STATUS_TIME")
                                excelWorksheet.Cells(iRowPosition, B_FINAL_CLEARED_DATE_COLUMN) = acClearedTime
                            End If
                        End If
                        adoStatusTime.Close
                        
                        excelWorksheet.Cells(iRowPosition, B_FINAL_CREATE_TO_CLEARED_COLUMN) = "=+(RC[-" + Trim(Str(B_FINAL_CREATE_TO_CLEARED_COLUMN - B_FINAL_CLEARED_DATE_COLUMN)) + "]-RC[-" + Trim(Str(B_FINAL_CREATE_TO_CLEARED_COLUMN - B_CREATED_DATE_COLUMN)) + "])*1440"
                        excelWorksheet.Cells(iRowPosition, B_FINAL_RECEIPT_TO_ARRIVE_COLUMN) = "=+(RC[-" + Trim(Str(B_FINAL_RECEIPT_TO_ARRIVE_COLUMN - B_FINAL_ARRIVE_DATE_COLUMN)) + "]-RC[-" + Trim(Str(B_FINAL_RECEIPT_TO_ARRIVE_COLUMN - B_FINAL_RECEIPT_DATE_COLUMN)) + "])*1440"
                        excelWorksheet.Cells(iRowPosition, B_FINAL_ARRIVE_TO_CLEARED_COLUMN) = "=+(RC[-" + Trim(Str(B_FINAL_ARRIVE_TO_CLEARED_COLUMN - B_FINAL_CLEARED_DATE_COLUMN)) + "]-RC[-" + Trim(Str(B_FINAL_ARRIVE_TO_CLEARED_COLUMN - B_FINAL_ARRIVE_DATE_COLUMN)) + "])*1440"
                        excelWorksheet.Cells(iRowPosition, B_FINAL_ENROUTE_TO_ARRIVED_COLUMN) = "=+(RC[-" + Trim(Str(B_FINAL_ENROUTE_TO_ARRIVED_COLUMN - B_FINAL_ARRIVE_DATE_COLUMN)) + "]-RC[-" + Trim(Str(B_FINAL_ENROUTE_TO_ARRIVED_COLUMN - B_FINAL_ENROUTE_DATE_COLUMN)) + "])*1440"
                    End If
                    adoReferUnit.Close
                    
                Else
                    excelWorksheet.Cells(iRowPosition, A_CREATE_TO_ASSIGN_COLUMN) = "=+(RC[-" + Trim(Str(A_CREATE_TO_ASSIGN_COLUMN - A_DISPATCH_DATE_COLUMN)) + "]-RC[-" + Trim(Str(A_CREATE_TO_ASSIGN_COLUMN - A_CREATED_DATE_COLUMN)) + "])*1440"
                    excelWorksheet.Cells(iRowPosition, A_ASSIGNED_TO_ARRIVED_COLUMN) = "=+(RC[-" + Trim(Str(A_ASSIGNED_TO_ARRIVED_COLUMN - A_ARRIVED_DATE_COLUMN)) + "]-RC[-" + Trim(Str(A_ASSIGNED_TO_ARRIVED_COLUMN - A_DISPATCH_DATE_COLUMN)) + "])*1440"
                    excelWorksheet.Cells(iRowPosition, A_CREATE_TO_CLEARED_COLUMN) = "=+(RC[-" + Trim(Str(A_CREATE_TO_CLEARED_COLUMN - A_CLEARED_DATE_COLUMN)) + "]-RC[-" + Trim(Str(A_CREATE_TO_CLEARED_COLUMN - A_CREATED_DATE_COLUMN)) + "])*1440"
                    excelWorksheet.Cells(iRowPosition, A_CREATE_TO_ARRIVED_COLUMN) = "=+(RC[-" + Trim(Str(A_CREATE_TO_ARRIVED_COLUMN - A_ARRIVED_DATE_COLUMN)) + "]-RC[-" + Trim(Str(A_CREATE_TO_ARRIVED_COLUMN - A_CREATED_DATE_COLUMN)) + "])*1440"
                    excelWorksheet.Cells(iRowPosition, A_ENROUTE_TO_ARRIVED_COLUMN) = "=+(RC[-" + Trim(Str(A_ENROUTE_TO_ARRIVED_COLUMN - A_ARRIVED_DATE_COLUMN)) + "]-RC[-" + Trim(Str(A_ENROUTE_TO_ARRIVED_COLUMN - A_ENROUTE_DATE_COLUMN)) + "])*1440"
                    excelWorksheet.Cells(iRowPosition, A_ENROUTE_TO_CLEARED_COLUMN) = "=+(RC[-" + Trim(Str(A_ENROUTE_TO_CLEARED_COLUMN - A_CLEARED_DATE_COLUMN)) + "]-RC[-" + Trim(Str(A_ENROUTE_TO_CLEARED_COLUMN - A_ENROUTE_DATE_COLUMN)) + "])*1440"
                    excelWorksheet.Cells(iRowPosition, A_ARRIVED_TO_CLEARED_COLUMN) = "=+(RC[-" + Trim(Str(A_ARRIVED_TO_CLEARED_COLUMN - A_CLEARED_DATE_COLUMN)) + "]-RC[-" + Trim(Str(A_ARRIVED_TO_CLEARED_COLUMN - A_ARRIVED_DATE_COLUMN)) + "])*1440"
                End If
                
                iRowPosition = iRowPosition + 1
                iCrewTotal = iCrewTotal + 1
                If adoUnitList.Fields("CALL_TYPE_CD") = "ELEC" Then
                    If blPriority1 Then
                        iPriorityPosition = iPriorityPosition + 1
                    Else
                        iPriority2Position = iPriority2Position + 1
                    End If
                End If
                'Move to the next record.
                adoUnitList.MoveNext
                
                'Update the progress bar
                i = i + 1
                frmMain.ProgressBar1.Value = i + 1
                frmMain.sbStatusBar.Panels(1).Text = "Processing " + Trim(excelWorksheet.Name) + " for unit : " + acPreviousUnitNo
            Loop
        
            '* Resize all the columns to fit the data
            For x = 1 To 3
                Set excelWorksheet = excelWorkBook.Worksheets(x)
                excelWorksheet.Select
                excelExcelApp.Cells.Select
                excelExcelApp.Selection.Columns.AutoFit
                excelExcelApp.Range("A2").Select
            Next
            
        End If
    End If
    frmMain.ProgressBar1.Value = 1
    frmMain.sbStatusBar.Panels(1).Text = "Saving excel file... "
    
    adoUnitList.Close
                
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

    If iHeaderType = 1 Then
        excelWorksheet.Cells(1, A_CAD_ID_COLUMN) = "CAD-ID"
        excelWorksheet.Cells(1, A_UNIT_COLUMN) = "Unit"
        excelWorksheet.Cells(1, A_PROBLEM_CODE_COLUMN) = "Problem Code"
        excelWorksheet.Cells(1, A_REFER_TO_COLUMN) = "Refer To"
        excelWorksheet.Cells(1, A_CREATED_DATE_COLUMN) = "Created"
        excelWorksheet.Columns(A_CREATED_DATE_COLUMN).Select
        excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
        
        excelWorksheet.Cells(1, A_DISPATCH_DATE_COLUMN) = "Dispatched"
        excelWorksheet.Columns(A_DISPATCH_DATE_COLUMN).Select
        excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
        
        excelWorksheet.Cells(1, A_ACCEPTED_DATE_COLUMN) = "Accepted"
        excelWorksheet.Columns(A_ACCEPTED_DATE_COLUMN).Select
        excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
        
        excelWorksheet.Cells(1, A_ENROUTE_DATE_COLUMN) = "EnRoute"
        excelWorksheet.Columns(A_ENROUTE_DATE_COLUMN).Select
        excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
        
        excelWorksheet.Cells(1, A_ARRIVED_DATE_COLUMN) = "Arrived"
        excelWorksheet.Columns(A_ARRIVED_DATE_COLUMN).Select
        excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
        
        excelWorksheet.Cells(1, A_REPORT_DATE_COLUMN) = "Reporting"
        excelWorksheet.Columns(A_REPORT_DATE_COLUMN).Select
        excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
        
        excelWorksheet.Cells(1, A_CLEARED_DATE_COLUMN) = "Cleared"
        excelWorksheet.Columns(A_CLEARED_DATE_COLUMN).Select
        excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
        
        excelWorksheet.Cells(1, A_CREATE_TO_ASSIGN_COLUMN) = "Create-Assign"
        excelWorksheet.Columns(A_CREATE_TO_ASSIGN_COLUMN).Select
        excelApplication.Selection.NumberFormat = "0"
        
        excelWorksheet.Cells(1, A_ASSIGNED_TO_ARRIVED_COLUMN) = "Assigned-Arrived"
        excelWorksheet.Columns(A_ASSIGNED_TO_ARRIVED_COLUMN).Select
        excelApplication.Selection.NumberFormat = "0"
        
        excelWorksheet.Cells(1, A_CREATE_TO_CLEARED_COLUMN) = "Create-Cleared"
        excelWorksheet.Columns(A_CREATE_TO_CLEARED_COLUMN).Select
        excelApplication.Selection.NumberFormat = "0"
        
        excelWorksheet.Cells(1, A_CREATE_TO_ARRIVED_COLUMN) = "Create-Arrived"
        excelWorksheet.Columns(A_CREATE_TO_ARRIVED_COLUMN).Select
        excelApplication.Selection.NumberFormat = "0"
        
        excelWorksheet.Cells(1, A_ENROUTE_TO_ARRIVED_COLUMN) = "EnRoute-Arrived"
        excelWorksheet.Columns(A_ENROUTE_TO_ARRIVED_COLUMN).Select
        excelApplication.Selection.NumberFormat = "0"
        
        excelWorksheet.Cells(1, A_ENROUTE_TO_CLEARED_COLUMN) = "EnRoute-Cleared"
        excelWorksheet.Columns(A_ENROUTE_TO_CLEARED_COLUMN).Select
        excelApplication.Selection.NumberFormat = "0"
        
        excelWorksheet.Cells(1, A_ARRIVED_TO_CLEARED_COLUMN) = "Arrived-Cleared"
        excelWorksheet.Columns(A_ARRIVED_TO_CLEARED_COLUMN).Select
        excelApplication.Selection.NumberFormat = "0"
    
        excelWorksheet.Cells(1, A_TOTAL_CALLS_COLUMN) = "Total Calls"
        excelWorksheet.Columns(A_TOTAL_CALLS_COLUMN).Select
        excelApplication.Selection.NumberFormat = "0"
        
        excelWorksheet.Cells(1, A_TOTAL_CUSTOMERS_COLUMN) = "Total Customers"
        excelWorksheet.Columns(A_TOTAL_CUSTOMERS_COLUMN).Select
        excelApplication.Selection.NumberFormat = "0"
    Else
        excelWorksheet.Range("A1", "B1").Select
        excelApplication.Selection.Interior.ColorIndex = 6
        excelWorksheet.Range("D1", "D1").Select
        excelApplication.Selection.Interior.ColorIndex = 6
        excelWorksheet.Range("M1", "M1").Select
        excelApplication.Selection.Interior.ColorIndex = 6
        excelWorksheet.Range("X1", "Z1").Select
        excelApplication.Selection.Interior.ColorIndex = 6
        excelWorksheet.Range("C1", "C1").Select
        excelApplication.Selection.Interior.ColorIndex = 4
        excelWorksheet.Range("E1", "L1").Select
        excelApplication.Selection.Interior.ColorIndex = 4
        excelWorksheet.Range("N1", "W1").Select
        excelApplication.Selection.Interior.ColorIndex = 45
        
        excelWorksheet.Cells(1, B_RESTORED_BY_SOCC) = "Restored by SOCC"
        excelWorksheet.Cells(1, B_CAD_ID_COLUMN) = "CAD-ID"
        excelWorksheet.Cells(1, B_PROBLEM_CODE_COLUMN) = "Job Type"
        excelWorksheet.Cells(1, B_INITIAL_UNIT_COLUMN) = "Initial Crew (Station 2)"
        
        excelWorksheet.Cells(1, B_CREATED_DATE_COLUMN) = "Created"
        excelWorksheet.Columns(B_CREATED_DATE_COLUMN).Select
        excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
        
        excelWorksheet.Cells(1, B_INITIAL_ENROUTE_DATE_COLUMN) = "EnRoute"
        excelWorksheet.Columns(B_INITIAL_ENROUTE_DATE_COLUMN).Select
        excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
        
        excelWorksheet.Cells(1, B_INITIAL_ARRIVED_DATE_COLUMN) = "Arrived"
        excelWorksheet.Columns(B_INITIAL_ARRIVED_DATE_COLUMN).Select
        excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
        
        excelWorksheet.Cells(1, B_INITIAL_CLEARED_DATE_COLUMN) = "Cleared"
        excelWorksheet.Columns(B_INITIAL_CLEARED_DATE_COLUMN).Select
        excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
         
        excelWorksheet.Cells(1, B_INITIAL_CREATE_TO_CLEARED_COLUMN) = "Create-Clear"
        excelWorksheet.Columns(B_INITIAL_CREATE_TO_CLEARED_COLUMN).Select
        excelApplication.Selection.NumberFormat = "0"
        
        excelWorksheet.Cells(1, B_INITIAL_CREATE_TO_ARRIVED_COLUMN) = "Create-Arrive"
        excelWorksheet.Columns(B_INITIAL_CREATE_TO_ARRIVED_COLUMN).Select
        excelApplication.Selection.NumberFormat = "0"
        
        excelWorksheet.Cells(1, B_INITIAL_ARRIVED_TO_CLEARED_COLUMN) = "Arrive-Cleared"
        excelWorksheet.Columns(B_INITIAL_ARRIVED_TO_CLEARED_COLUMN).Select
        excelApplication.Selection.NumberFormat = "0"
        
        excelWorksheet.Cells(1, B_INITIAL_ENROUTE_TO_ARRIVED_COLUMN) = "Enroute-Arrive"
        excelWorksheet.Columns(B_INITIAL_ENROUTE_TO_ARRIVED_COLUMN).Select
        excelApplication.Selection.NumberFormat = "0"
        
        excelWorksheet.Cells(1, B_RESTORE_DATE_COLUMN) = "Service Restored"
        excelWorksheet.Columns(B_RESTORE_DATE_COLUMN).Select
        excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
        
        excelWorksheet.Cells(1, B_REFER_TO_COLUMN) = "Referred"
        excelWorksheet.Cells(1, B_FINAL_UNIT_COLUMN) = "Referred Crew"
        
        excelWorksheet.Cells(1, B_FINAL_RECEIPT_DATE_COLUMN) = "Ticket Receipt"
        excelWorksheet.Columns(B_FINAL_RECEIPT_DATE_COLUMN).Select
        excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
        
        excelWorksheet.Cells(1, B_FINAL_ENROUTE_DATE_COLUMN) = "Enroute"
        excelWorksheet.Columns(B_FINAL_ENROUTE_DATE_COLUMN).Select
        excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
        
        excelWorksheet.Cells(1, B_FINAL_ARRIVE_DATE_COLUMN) = "Arrive"
        excelWorksheet.Columns(B_FINAL_ARRIVE_DATE_COLUMN).Select
        excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
        
        excelWorksheet.Cells(1, B_FINAL_CLEARED_DATE_COLUMN) = "Cleared"
        excelWorksheet.Columns(B_FINAL_CLEARED_DATE_COLUMN).Select
        excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
        
        excelWorksheet.Cells(1, B_FINAL_CREATE_TO_CLEARED_COLUMN) = "Create-Cleared"
        excelWorksheet.Columns(B_FINAL_CREATE_TO_CLEARED_COLUMN).Select
        excelApplication.Selection.NumberFormat = "0"
        
        excelWorksheet.Cells(1, B_FINAL_RECEIPT_TO_ARRIVE_COLUMN) = "Receipt-Arrive"
        excelWorksheet.Columns(B_FINAL_RECEIPT_TO_ARRIVE_COLUMN).Select
        excelApplication.Selection.NumberFormat = "0"
        
        excelWorksheet.Cells(1, B_FINAL_ARRIVE_TO_CLEARED_COLUMN) = "Arrive-Clear"
        excelWorksheet.Columns(B_FINAL_ARRIVE_TO_CLEARED_COLUMN).Select
        excelApplication.Selection.NumberFormat = "0"
        
        excelWorksheet.Cells(1, B_FINAL_ENROUTE_TO_ARRIVED_COLUMN) = "Enroute-Arrive"
        excelWorksheet.Columns(B_FINAL_ENROUTE_TO_ARRIVED_COLUMN).Select
        excelApplication.Selection.NumberFormat = "0"
        
        excelWorksheet.Cells(1, B_FINAL_CREATE_TO_RESTORE_COLUMN) = "Create-Restore"
        excelWorksheet.Columns(B_FINAL_CREATE_TO_RESTORE_COLUMN).Select
        excelApplication.Selection.NumberFormat = "0"
        
        excelWorksheet.Cells(1, B_TOTAL_CALLS_COLUMN) = "Total Calls"
        excelWorksheet.Columns(B_TOTAL_CALLS_COLUMN).Select
        excelApplication.Selection.NumberFormat = "0"
        
        excelWorksheet.Cells(1, B_TOTAL_CUSTOMERS_COLUMN) = "Total Customers"
        excelWorksheet.Columns(B_TOTAL_CUSTOMERS_COLUMN).Select
        excelApplication.Selection.NumberFormat = "0"
    End If
    
End Sub




