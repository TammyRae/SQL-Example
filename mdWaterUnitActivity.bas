Attribute VB_Name = "mdWaterUnitActivity"
Option Explicit

'******************************************
'*** Excel column constants for report ****
'******************************************
Public Const CAD_ID_COLUMN As Integer = 1
Public Const UNIT_COLUMN As Integer = 2
Public Const PROBLEM_CODE_COLUMN As Integer = 3
Public Const REFER_TO_COLUMN As Integer = 4
Public Const CREATED_DATE_COLUMN As Integer = 5
Public Const DISPATCH_DATE_COLUMN As Integer = 6
Public Const ACCEPTED_DATE_COLUMN As Integer = 7
Public Const ENROUTE_DATE_COLUMN As Integer = 8
Public Const ARRIVED_DATE_COLUMN As Integer = 9
Public Const REPORT_DATE_COLUMN As Integer = 10
Public Const CLEARED_DATE_COLUMN As Integer = 11
Public Const CREATE_TO_ASSIGN_COLUMN As Integer = 12
Public Const ASSIGNED_TO_ARRIVED_COLUMN As Integer = 13
Public Const CREATE_TO_CLEARED_COLUMN As Integer = 14
Public Const CREATE_TO_ARRIVED_COLUMN As Integer = 15
Public Const ENROUTE_TO_ARRIVED_COLUMN As Integer = 16
Public Const ENROUTE_TO_CLEARED_COLUMN As Integer = 17
Public Const ARRIVED_TO_CLEARED_COLUMN As Integer = 18
Public Const WORK_DAY = 19
Public Const WORK_TIME = 20
Public Const DAY_OF_WEEK = 21


Rem ****************************************************************************
Rem *
Rem * Function Name:    GenerateWaterUnitActivityReport
Rem * Paramters    :    acUserName      - User name to login to cad database
Rem *                   acPassword      - Password to use to login to database.
Rem *                   acDBInstance    - The CAD db instance
Rem *                   acOutputName    - File name to same output (spreadsheet) to.
Rem *
Rem * Description   : Generate the unit activity report
Rem ****************************************************************************
Sub GenerateWaterUnitActivityReport(acUserName As String, acPassword As String, _
                                    acDBInstance As String, acExcelFile As String, acReportType As String)
    Dim acSQLText As String, acIsrNo As String, acUnitNo As String
    Dim acCellData As String, acPreviousUnitNo As String
    Dim acDispatchTime As String, acReportingTime As String
    Dim acAcceptedTime As String, acEnrouteTime As String
    Dim acArrivedTime As String, acClearedTime As String, acAgencyCode As String
    Dim acCreatedTime As String, acProbCode As String, acRefer As String
    Dim acWorkDay As String, acWorkTime As String, acDayOfWeek As String
    Dim i As Integer, x As Integer, iTotalRows As Double, iRowPosition As Double
    Dim iRow1Pos As Integer, iRow2Pos As Integer, iRow3Pos As Integer, iRow4Pos As Integer
    Dim iRow5Pos As Integer, iRow6Pos As Integer, iRow7Pos As Integer, iRow8Pos As Integer
    Dim iRow9Pos As Integer, iRow10Pos As Integer, iRow11Pos As Integer, blIsPriority As Boolean
    Dim iCounter As Integer, blPrintRow As Boolean
    Dim iCrewTotal As Integer, iTimeDifference As Integer
    Dim excelExcelApp As Excel.Application ' This is the excel program
    Dim excelWorkBook As Excel.Workbook ' This is the work book
    Dim excelWorksheet As Excel.Worksheet ' This is the sheet
    Dim adoConnection As ADODB.Connection
    Dim adoUnitList As ADODB.Recordset
    Dim adoStatusTime As ADODB.Recordset
    
    Set adoConnection = New ADODB.Connection
    Set adoUnitList = New ADODB.Recordset
    Set adoStatusTime = New ADODB.Recordset
  
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
    
    acSQLText = " select /*+ INDEX (d11_unit_activity J_D11_UNIT_ACTIV_5) */ ISR.ISR_NO, UNIT, ISR.PRIORITY, P_FILTER1, "
    acSQLText = acSQLText + " DECODE(ISR.init_service_code, 'TDPM', 'EROU', (SELECT CALL_TYPE_CD FROM CLUE@fmsrpt_ompr11_l WHERE CLUE_CD = ISR.INIT_SERVICE_CODE)) CALL_TYPE_CD,  "
    acSQLText = acSQLText + " ISR.init_service_code, ISR.AGENCY_CODE, "
    acSQLText = acSQLText + " TO_CHAR( TO_DATE(D11_START_DATE || D11_START_TIME, 'YYYYMMDDHH24MISS'), 'MM/DD/YYYY HH24:MI:SS') FORMATED_START_TIME, "

    acSQLText = acSQLText + " to_date(D11_START_DATE, 'YYYYMMDD') WORK_DATE, "
    acSQLText = acSQLText + " to_char(to_date(D11_START_TIME, 'HH24MISS'), 'HH24:MI') WORK_TIME, "
    acSQLText = acSQLText + " to_char(to_date(D11_START_DATE || D11_START_TIME, 'YYYYMMDDHH24MISS'), 'DAY') DAY_OF_WEEK, "
    
    acSQLText = acSQLText + " TO_CHAR( TO_DATE(ISR.CREATED_DATE || ISR.CREATED_TIME, 'YYYYMMDDHH24MISS'), 'MM/DD/YYYY HH24:MI:SS') CREATION_TIME, ISR.UDF11 TOTAL_CALLS, ISR.UDF7 CUSTOMERS_AFFECTED "
    acSQLText = acSQLText + " From d11_unit_activity, ISR "
    acSQLText = acSQLText + " Where ISR.ISR_NO = d11_unit_activity.ISR_NO "
    acSQLText = acSQLText + " and b07_status_code = 'RP' "
    acSQLText = acSQLText + " and d11_start_date >= '" + Format(frmMain.lstDatePickerFrom.Value, "YYYYMMDD") + "' "
    acSQLText = acSQLText + " and d11_end_date <= '" + Format(frmMain.lstDatePickerTo.Value, "YYYYMMDD") + "' "
    acSQLText = acSQLText + " and ISR.agency_Code in ('WTR', 'SWR') "
    acSQLText = acSQLText + " and ((length(d11_unit_activity.unit) = 3 and d11_unit_activity.unit > '199' and d11_unit_activity.unit < '300') or "
    acSQLText = acSQLText + " (length(d11_unit_activity.unit) = 3 and d11_unit_activity.unit > '399' and d11_unit_activity.unit < '500') or "
    acSQLText = acSQLText + " (length(d11_unit_activity.unit) = 4 and d11_unit_activity.unit > '3300' and d11_unit_activity.unit < '3400')) "
    acSQLText = acSQLText + " order by CALL_TYPE_CD, UNIT, D11_START_DATE, D11_START_TIME "

    frmMain.sbStatusBar.Panels(1).Text = "Running query..."
    adoUnitList.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText

    If Not (adoUnitList.BOF And adoUnitList.EOF) Then
        iTotalRows = adoUnitList.RecordCount
          
        If iTotalRows > 0 Then
            For iCounter = excelWorkBook.Worksheets.Count To 10  '** There is usually 3 worksheets open by default in excel
                excelWorkBook.Worksheets.Add
            Next
            
            ' Create each worksheet and the column headers.
            Set excelWorksheet = excelWorkBook.Worksheets(1)
            excelWorksheet.Name = "Total"
            excelWorksheet.Select
            CreateWorksheetHeader excelWorksheet, excelWorkBook, excelExcelApp
            Set excelWorksheet = excelWorkBook.Worksheets(2)
            excelWorksheet.Name = "TotSocc"
            excelWorksheet.Select
            CreateWorksheetHeader excelWorksheet, excelWorkBook, excelExcelApp
            Set excelWorksheet = excelWorkBook.Worksheets(3)
            excelWorksheet.Select
            excelWorksheet.Name = "TotRid"
            CreateWorksheetHeader excelWorksheet, excelWorkBook, excelExcelApp
            Set excelWorksheet = excelWorkBook.Worksheets(4)
            excelWorksheet.Select
            excelWorksheet.Name = "SOCCPrty"
            CreateWorksheetHeader excelWorksheet, excelWorkBook, excelExcelApp
            Set excelWorksheet = excelWorkBook.Worksheets(5)
            excelWorksheet.Select
            excelWorksheet.Name = "SOCCRtne"
            CreateWorksheetHeader excelWorksheet, excelWorkBook, excelExcelApp
            Set excelWorksheet = excelWorkBook.Worksheets(6)
            excelWorksheet.Select
            excelWorksheet.Name = "WtrPrty"
            CreateWorksheetHeader excelWorksheet, excelWorkBook, excelExcelApp
            Set excelWorksheet = excelWorkBook.Worksheets(7)
            excelWorksheet.Select
            excelWorksheet.Name = "WtrRtne"
            CreateWorksheetHeader excelWorksheet, excelWorkBook, excelExcelApp
            Set excelWorksheet = excelWorkBook.Worksheets(8)
            excelWorksheet.Select
            excelWorksheet.Name = "SwrPrty"
            CreateWorksheetHeader excelWorksheet, excelWorkBook, excelExcelApp
            Set excelWorksheet = excelWorkBook.Worksheets(9)
            excelWorksheet.Select
            excelWorksheet.Name = "SwrRtne"
            CreateWorksheetHeader excelWorksheet, excelWorkBook, excelExcelApp
            Set excelWorksheet = excelWorkBook.Worksheets(10)
            excelWorksheet.Select
            excelWorksheet.Name = "LPSP"
            CreateWorksheetHeader excelWorksheet, excelWorkBook, excelExcelApp
            Set excelWorksheet = excelWorkBook.Worksheets(11)
            excelWorksheet.Select
            excelWorksheet.Name = "Misc"
            CreateWorksheetHeader excelWorksheet, excelWorkBook, excelExcelApp
           
            'Update the progress bar
            frmMain.ProgressBar1.Max = iTotalRows + 1
            frmMain.ProgressBar1.Min = 1
            frmMain.ProgressBar1.Visible = True
            
            i = 0
            adoUnitList.MoveFirst
            acPreviousUnitNo = adoUnitList.Fields("UNIT")
            iCrewTotal = 0
            iRow1Pos = 2
            iRow2Pos = 2
            iRow3Pos = 2
            iRow4Pos = 2
            iRow5Pos = 2
            iRow6Pos = 2
            iRow7Pos = 2
            iRow8Pos = 2
            iRow9Pos = 2
            iRow10Pos = 2
            iRow11Pos = 2
            
            Set excelWorksheet = excelWorkBook.Worksheets(1)
            Do While Not adoUnitList.EOF
                acIsrNo = adoUnitList.Fields.Item("ISR_NO")
                acReportingTime = adoUnitList.Fields("FORMATED_START_TIME")
                acCreatedTime = adoUnitList.Fields("CREATION_TIME")
                acUnitNo = adoUnitList.Fields("UNIT")
                acRefer = adoUnitList.Fields("P_FILTER1")
                acProbCode = adoUnitList.Fields("INIT_SERVICE_CODE")
                acAgencyCode = adoUnitList.Fields("AGENCY_CODE")
                
                acWorkDay = adoUnitList.Fields("WORK_DATE")
                acWorkTime = adoUnitList.Fields("WORK_TIME")
                acDayOfWeek = adoUnitList.Fields("DAY_OF_WEEK")
                
                blIsPriority = InStr(1, "BKPH COFF CON JTLN MHOF NW PSA SBRK LPSP SRLK SSO WBRK", acProbCode)
                
                acDispatchTime = ""
                acAcceptedTime = ""
                acEnrouteTime = ""
                acArrivedTime = ""
                acClearedTime = ""
                
                '********************** Get the dispatch time *********************
                acSQLText = "select to_char(max(to_date(d11_start_date || d11_start_time, 'YYYYMMDDHH24MISS')), 'MM/DD/YYYY HH24:MI:SS') STATUS_TIME "
                acSQLText = acSQLText + " from d11_unit_activity "
                acSQLText = acSQLText + " where isr_no = '" + acIsrNo + "' "
                acSQLText = acSQLText + " and b07_status_code = 'DP' and unit = '" + acUnitNo + "' "
                acSQLText = acSQLText + " and (d11_start_date || d11_start_time) <= '" + Mid(acReportingTime, 7, 4) + Mid(acReportingTime, 1, 2) + Mid(acReportingTime, 4, 2) + Mid(acReportingTime, 12, 2) + Mid(acReportingTime, 15, 2) + Mid(acReportingTime, 18, 2) + "' "
                    
                adoStatusTime.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                If Not (adoStatusTime.BOF And adoStatusTime.EOF) Then
                   If Not IsNull(adoStatusTime.Fields("STATUS_TIME")) Then
                       acDispatchTime = adoStatusTime.Fields("STATUS_TIME")
                   End If
                End If
                adoStatusTime.Close
                
                '********************** Get the accepted time *********************
                acSQLText = "select to_char(max(to_date(d11_start_date || d11_start_time, 'YYYYMMDDHH24MISS')), 'MM/DD/YYYY HH24:MI') STATUS_TIME "
                acSQLText = acSQLText + " from d11_unit_activity "
                acSQLText = acSQLText + " where isr_no = '" + acIsrNo + "' "
                acSQLText = acSQLText + " and b07_status_code = 'AC' and unit = '" + acUnitNo + "' "
                acSQLText = acSQLText + " and (d11_start_date || d11_start_time) <= '" + Mid(acReportingTime, 7, 4) + Mid(acReportingTime, 1, 2) + Mid(acReportingTime, 4, 2) + Mid(acReportingTime, 12, 2) + Mid(acReportingTime, 15, 2) + Mid(acReportingTime, 18, 2) + "' "
                    
                adoStatusTime.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                If Not (adoStatusTime.BOF And adoStatusTime.EOF) Then
                   If Not IsNull(adoStatusTime.Fields("STATUS_TIME")) Then
                       acAcceptedTime = adoStatusTime.Fields("STATUS_TIME")
                   End If
                End If
                adoStatusTime.Close
                
                '********************** Get the enroute time *********************
                acSQLText = "select to_char(max(to_date(d11_start_date || d11_start_time, 'YYYYMMDDHH24MISS')), 'MM/DD/YYYY HH24:MI') STATUS_TIME "
                acSQLText = acSQLText + " from d11_unit_activity "
                acSQLText = acSQLText + " where isr_no = '" + acIsrNo + "' "
                acSQLText = acSQLText + " and b07_status_code = 'ER' and unit = '" + acUnitNo + "' "
                acSQLText = acSQLText + " and (d11_start_date || d11_start_time) <= '" + Mid(acReportingTime, 7, 4) + Mid(acReportingTime, 1, 2) + Mid(acReportingTime, 4, 2) + Mid(acReportingTime, 12, 2) + Mid(acReportingTime, 15, 2) + Mid(acReportingTime, 18, 2) + "' "
                     
                adoStatusTime.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                If Not (adoStatusTime.BOF And adoStatusTime.EOF) Then
                    If Not IsNull(adoStatusTime.Fields("STATUS_TIME")) Then
                        acEnrouteTime = adoStatusTime.Fields("STATUS_TIME")
                    End If
                End If
                adoStatusTime.Close
                
                '********************** Get the arrived time *********************
                acSQLText = "select to_char(max(to_date(d11_start_date || d11_start_time, 'YYYYMMDDHH24MISS')), 'MM/DD/YYYY HH24:MI') STATUS_TIME "
                acSQLText = acSQLText + " from d11_unit_activity "
                acSQLText = acSQLText + " where isr_no = '" + acIsrNo + "' "
                acSQLText = acSQLText + " and b07_status_code = 'AR' and unit = '" + acUnitNo + "' "
                acSQLText = acSQLText + " and (d11_start_date || d11_start_time) <= '" + Mid(acReportingTime, 7, 4) + Mid(acReportingTime, 1, 2) + Mid(acReportingTime, 4, 2) + Mid(acReportingTime, 12, 2) + Mid(acReportingTime, 15, 2) + Mid(acReportingTime, 18, 2) + "' "
                     
                adoStatusTime.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                If Not (adoStatusTime.BOF And adoStatusTime.EOF) Then
                    If Not IsNull(adoStatusTime.Fields("STATUS_TIME")) Then
                        acArrivedTime = adoStatusTime.Fields("STATUS_TIME")
                    End If
                End If
                adoStatusTime.Close
                
                '********************** Get the cleared time *********************
                acSQLText = "select to_char(min(TO_DATE(c07_start_date || c07_start_time, 'YYYYMMDDHH24MISS')), 'MM/DD/YYYY HH24:MI') STATUS_TIME "
                acSQLText = acSQLText + " from c07_isr_activity "
                acSQLText = acSQLText + " where isr_no = '" + acIsrNo + "' "
                acSQLText = acSQLText + " and c07_activity_code = 'CL' "
                acSQLText = acSQLText + " and (c07_start_date || c07_start_time) >= '" + Mid(acReportingTime, 7, 4) + Mid(acReportingTime, 1, 2) + Mid(acReportingTime, 4, 2) + Mid(acReportingTime, 12, 2) + Mid(acReportingTime, 15, 2) + Mid(acReportingTime, 18, 2) + "' "
                
                adoStatusTime.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText
                If Not (adoStatusTime.BOF And adoStatusTime.EOF) Then
                    If Not IsNull(adoStatusTime.Fields("STATUS_TIME")) Then
                        acClearedTime = adoStatusTime.Fields("STATUS_TIME")
                    End If
                End If
                adoStatusTime.Close
                         
                For iCounter = 1 To 5
                    blPrintRow = False
                    Select Case iCounter
                        Case 1:
                            iRowPosition = iRow1Pos
                            iRow1Pos = iRow1Pos + 1
                            blPrintRow = True
                            Set excelWorksheet = excelWorkBook.Worksheets(1)
                        Case 2:
                            If Trim(acUnitNo) = "200" Or Mid(acUnitNo, 1, 1) = "4" Then  '** 400 crews are Ridenour
                                iRowPosition = iRow3Pos
                                iRow3Pos = iRow3Pos + 1    '** Row 3 is the total-ridenour
                                Set excelWorksheet = excelWorkBook.Worksheets(3)
                            Else
                                iRowPosition = iRow2Pos
                                iRow2Pos = iRow2Pos + 1    '** Row 2 is the total-socc
                                Set excelWorksheet = excelWorkBook.Worksheets(2)
                            End If
                            blPrintRow = True
                            
                        Case 3:
                            If Trim(acUnitNo) = "200" Or Mid(acUnitNo, 1, 1) = "4" Then
                                blPrintRow = False
                            Else
                                If blIsPriority = True Then
                                    iRowPosition = iRow4Pos
                                    iRow4Pos = iRow4Pos + 1
                                    Set excelWorksheet = excelWorkBook.Worksheets(4)
                                Else
                                    iRowPosition = iRow5Pos
                                    iRow5Pos = iRow5Pos + 1
                                    Set excelWorksheet = excelWorkBook.Worksheets(5)
                                End If
                                blPrintRow = True
                            End If
                            
                        Case 4:
                            If Trim(acUnitNo) = "200" Or Mid(acUnitNo, 1, 1) = "4" Then
                                blPrintRow = False
                            Else
                                If Trim(acAgencyCode) = "WTR" Then
                                    If blIsPriority = True Then
                                        iRowPosition = iRow6Pos
                                        iRow6Pos = iRow6Pos + 1
                                        Set excelWorksheet = excelWorkBook.Worksheets(6)
                                    Else
                                        iRowPosition = iRow7Pos
                                        iRow7Pos = iRow7Pos + 1
                                        Set excelWorksheet = excelWorkBook.Worksheets(7)
                                    End If
                                Else
                                    If blIsPriority = True Then
                                        iRowPosition = iRow8Pos
                                        iRow8Pos = iRow8Pos + 1
                                        Set excelWorksheet = excelWorkBook.Worksheets(8)
                                    Else
                                        iRowPosition = iRow9Pos
                                        iRow9Pos = iRow9Pos + 1
                                        Set excelWorksheet = excelWorkBook.Worksheets(9)
                                    End If
                                End If
                                blPrintRow = True
                            End If
                        Case 5:
                            If Trim(acUnitNo) = "200" Or Mid(acUnitNo, 1, 1) = "4" Then
                                blPrintRow = False
                            Else
                                If Trim(acProbCode) = "LPSP" Then
                                    iRowPosition = iRow10Pos
                                    iRow10Pos = iRow10Pos + 1
                                    Set excelWorksheet = excelWorkBook.Worksheets(10)
                                    blPrintRow = True
                                End If
                            End If
                    End Select
                            
                    If blPrintRow Then
                        excelWorksheet.Cells(iRowPosition, CAD_ID_COLUMN) = acIsrNo
                        excelWorksheet.Cells(iRowPosition, UNIT_COLUMN) = acUnitNo
                        excelWorksheet.Cells(iRowPosition, PROBLEM_CODE_COLUMN) = acProbCode
                        excelWorksheet.Cells(iRowPosition, REFER_TO_COLUMN) = acRefer
                        excelWorksheet.Cells(iRowPosition, CREATED_DATE_COLUMN) = acCreatedTime
                        excelWorksheet.Cells(iRowPosition, REPORT_DATE_COLUMN) = acReportingTime
                        excelWorksheet.Cells(iRowPosition, DISPATCH_DATE_COLUMN) = acDispatchTime
                        excelWorksheet.Cells(iRowPosition, ACCEPTED_DATE_COLUMN) = acAcceptedTime
                        excelWorksheet.Cells(iRowPosition, ENROUTE_DATE_COLUMN) = acEnrouteTime
                        excelWorksheet.Cells(iRowPosition, ARRIVED_DATE_COLUMN) = acArrivedTime
                        excelWorksheet.Cells(iRowPosition, CLEARED_DATE_COLUMN) = acClearedTime
                        '******** Put the formulas for the time differences columns **********************
                        excelWorksheet.Cells(iRowPosition, CREATE_TO_ASSIGN_COLUMN) = "=+(RC[-" + Trim(Str(CREATE_TO_ASSIGN_COLUMN - DISPATCH_DATE_COLUMN)) + "]-RC[-" + Trim(Str(CREATE_TO_ASSIGN_COLUMN - CREATED_DATE_COLUMN)) + "])*1440"
                        excelWorksheet.Cells(iRowPosition, ASSIGNED_TO_ARRIVED_COLUMN) = "=+(RC[-" + Trim(Str(ASSIGNED_TO_ARRIVED_COLUMN - ARRIVED_DATE_COLUMN)) + "]-RC[-" + Trim(Str(ASSIGNED_TO_ARRIVED_COLUMN - DISPATCH_DATE_COLUMN)) + "])*1440"
                        excelWorksheet.Cells(iRowPosition, CREATE_TO_CLEARED_COLUMN) = "=+(RC[-" + Trim(Str(CREATE_TO_CLEARED_COLUMN - CLEARED_DATE_COLUMN)) + "]-RC[-" + Trim(Str(CREATE_TO_CLEARED_COLUMN - CREATED_DATE_COLUMN)) + "])*1440"
                        excelWorksheet.Cells(iRowPosition, CREATE_TO_ARRIVED_COLUMN) = "=+(RC[-" + Trim(Str(CREATE_TO_ARRIVED_COLUMN - ARRIVED_DATE_COLUMN)) + "]-RC[-" + Trim(Str(CREATE_TO_ARRIVED_COLUMN - CREATED_DATE_COLUMN)) + "])*1440"
                        excelWorksheet.Cells(iRowPosition, ENROUTE_TO_ARRIVED_COLUMN) = "=+(RC[-" + Trim(Str(ENROUTE_TO_ARRIVED_COLUMN - ARRIVED_DATE_COLUMN)) + "]-RC[-" + Trim(Str(ENROUTE_TO_ARRIVED_COLUMN - ENROUTE_DATE_COLUMN)) + "])*1440"
                        excelWorksheet.Cells(iRowPosition, ENROUTE_TO_CLEARED_COLUMN) = "=+(RC[-" + Trim(Str(ENROUTE_TO_CLEARED_COLUMN - CLEARED_DATE_COLUMN)) + "]-RC[-" + Trim(Str(ENROUTE_TO_CLEARED_COLUMN - ENROUTE_DATE_COLUMN)) + "])*1440"
                        excelWorksheet.Cells(iRowPosition, ARRIVED_TO_CLEARED_COLUMN) = "=+(RC[-" + Trim(Str(ARRIVED_TO_CLEARED_COLUMN - CLEARED_DATE_COLUMN)) + "]-RC[-" + Trim(Str(ARRIVED_TO_CLEARED_COLUMN - ARRIVED_DATE_COLUMN)) + "])*1440"
                        
                        '* Set work day work time and day of week
                        excelWorksheet.Cells(iRowPosition, WORK_DAY) = acWorkDay
                        excelWorksheet.Cells(iRowPosition, WORK_TIME) = acWorkTime
                        excelWorksheet.Cells(iRowPosition, DAY_OF_WEEK) = acDayOfWeek
                        
                    End If
                    
                Next
                     
                'Move to the next record.
                adoUnitList.MoveNext
                iCrewTotal = iCrewTotal + 1
                
                'Update the progress bar
                i = i + 1
                frmMain.ProgressBar1.Value = i + 1
                frmMain.sbStatusBar.Panels(1).Text = "Processing " + Trim(excelWorksheet.Name) + " for unit : " + acUnitNo
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
Private Sub CreateWorksheetHeader(excelWorksheet As Excel.Worksheet, excelWorkBook As Excel.Workbook, excelApplication As Excel.Application)

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
    excelApplication.Range("A2").Select
    excelApplication.ActiveWindow.FreezePanes = True
    excelApplication.Rows("1:1").Select
    excelApplication.Selection.Font.Bold = True

    excelWorksheet.Cells(1, CAD_ID_COLUMN) = "CAD-ID"
    excelWorksheet.Cells(1, UNIT_COLUMN) = "Unit"
    excelWorksheet.Cells(1, PROBLEM_CODE_COLUMN) = "Problem Code"
    excelWorksheet.Cells(1, REFER_TO_COLUMN) = "Refer To"
    excelWorksheet.Cells(1, CREATED_DATE_COLUMN) = "Created"
    excelWorksheet.Columns(CREATED_DATE_COLUMN).Select
    excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
    
    excelWorksheet.Cells(1, DISPATCH_DATE_COLUMN) = "Dispatched"
    excelWorksheet.Columns(DISPATCH_DATE_COLUMN).Select
    excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
    
    excelWorksheet.Cells(1, ACCEPTED_DATE_COLUMN) = "Accepted"
    excelWorksheet.Columns(ACCEPTED_DATE_COLUMN).Select
    excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
    
    excelWorksheet.Cells(1, ENROUTE_DATE_COLUMN) = "EnRoute"
    excelWorksheet.Columns(ENROUTE_DATE_COLUMN).Select
    excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
    
    excelWorksheet.Cells(1, ARRIVED_DATE_COLUMN) = "Arrived"
    excelWorksheet.Columns(ARRIVED_DATE_COLUMN).Select
    excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
    
    excelWorksheet.Cells(1, REPORT_DATE_COLUMN) = "Reporting"
    excelWorksheet.Columns(REPORT_DATE_COLUMN).Select
    excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
    
    excelWorksheet.Cells(1, CLEARED_DATE_COLUMN) = "Cleared"
    excelWorksheet.Columns(CLEARED_DATE_COLUMN).Select
    excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"
    
    excelWorksheet.Cells(1, CREATE_TO_ASSIGN_COLUMN) = "Create-Assign"
    excelWorksheet.Columns(CREATE_TO_ASSIGN_COLUMN).Select
    excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, ASSIGNED_TO_ARRIVED_COLUMN) = "Assigned-Arrived"
    excelWorksheet.Columns(ASSIGNED_TO_ARRIVED_COLUMN).Select
    excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, CREATE_TO_CLEARED_COLUMN) = "Create-Cleared"
    excelWorksheet.Columns(CREATE_TO_CLEARED_COLUMN).Select
    excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, CREATE_TO_ARRIVED_COLUMN) = "Create-Arrived"
    excelWorksheet.Columns(CREATE_TO_ARRIVED_COLUMN).Select
    excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, ENROUTE_TO_ARRIVED_COLUMN) = "EnRoute-Arrived"
    excelWorksheet.Columns(ENROUTE_TO_ARRIVED_COLUMN).Select
    excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, ENROUTE_TO_CLEARED_COLUMN) = "EnRoute-Cleared"
    excelWorksheet.Columns(ENROUTE_TO_CLEARED_COLUMN).Select
    excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, ARRIVED_TO_CLEARED_COLUMN) = "Arrived-Cleared"
    excelWorksheet.Columns(ARRIVED_TO_CLEARED_COLUMN).Select
    excelApplication.Selection.NumberFormat = "0"

    excelWorksheet.Cells(1, WORK_DAY) = "Work Date"
    excelWorksheet.Columns(WORK_DAY).Select
    excelApplication.Selection.NumberFormat = "m/d/yy h:mm;@"

    excelWorksheet.Cells(1, WORK_TIME) = "Work Time"
    excelWorksheet.Cells(1, DAY_OF_WEEK) = "Day"

End Sub




