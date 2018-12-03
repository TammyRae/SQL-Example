Attribute VB_Name = "mdServiceCenterUnitActivity"
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
Public Const TOTAL_CALLS_COLUMN As Integer = 19
Public Const TOTAL_CUSTOMERS_COLUMN As Integer = 20



Rem ****************************************************************************
Rem *
Rem * Function Name:    GenerateServiceCenterUnitActivityReport
Rem * Paramters    :    acUserName      - User name to login to cad database
Rem *                   acPassword      - Password to use to login to database.
Rem *                   acDBInstance    - The CAD db instance
Rem *                   acOutputName    - File name to same output (spreadsheet) to.
Rem *
Rem * Description   : Generate the unit activity report for service center crews.
Rem ****************************************************************************
Sub GenerateServiceCenterUnitActivityReport(acUserName As String, acPassword As String, _
                                            acDBInstance As String, acExcelFile As String, acReportType As String)
    Dim acSQLText As String, acIsrNo As String, acUnitNo As String
    Dim acCellData As String, acPreviousUnitNo As String
    Dim acDispatchTime As String, acReportingTime As String
    Dim acAcceptedTime As String, acEnrouteTime As String
    Dim acArrivedTime As String, acClearedTime As String
    Dim acCreatedTime As String
    Dim i As Integer, x As Integer, iTotalRows As Double, iRowPosition As Double
    Dim iCrewTotal As Integer, iTimeDifference As Integer, iStartRow As Integer
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
    
    acSQLText = " select ISR.ISR_NO, UNIT, ISR.PRIORITY, P_FILTER1, "
    'PAS - 2/5/2009 Added Jobtypes PCRM to Decode per Request: 130679
    acSQLText = acSQLText + " DECODE(ISR.init_service_code, 'MCC OFF', 'EROU', 'MCC', 'EROU', 'POLE-DIS', 'EROU', 'POLE-SAF', 'EROU', 'PL-REC', 'EROU', 'PCRM', 'EROU', 'M-DISREM', 'EROU', 'PL-DIS', 'EROU', 'POLE-REC', 'EROU', 'PL-SAF', 'EROU', 'TDL-ONLY', 'EROU', 'TDLPM', 'EROU', 'TLAT', 'EROU', 'TDPM', 'EROU', 'DPNP', 'EROU', 'DPSF', 'EROU', 'RCPP', 'EROU', 'DPMI', 'EROU', (SELECT CALL_TYPE_CD FROM CLUE@fmsrpt_ompr11_l WHERE CLUE_CD = ISR.INIT_SERVICE_CODE)) CALL_TYPE_CD,  "
    acSQLText = acSQLText + " ISR.init_service_code, D11_UNIT_ACTIVITY.AGENCY_CODE, "
    acSQLText = acSQLText + " TO_CHAR( TO_DATE(D11_START_DATE || D11_START_TIME, 'YYYYMMDDHH24MISS'), 'MM/DD/YYYY HH24:MI:SS') FORMATED_START_TIME, "
    acSQLText = acSQLText + " TO_CHAR( TO_DATE(ISR.CREATED_DATE || ISR.CREATED_TIME, 'YYYYMMDDHH24MISS'), 'MM/DD/YYYY HH24:MI:SS') CREATION_TIME, ISR.UDF11 TOTAL_CALLS, ISR.UDF7 CUSTOMERS_AFFECTED "
    acSQLText = acSQLText + " From d11_unit_activity, ISR "
    acSQLText = acSQLText + " Where ISR.ISR_NO = d11_unit_activity.ISR_NO "
    acSQLText = acSQLText + " and b07_status_code = 'RP' "
    acSQLText = acSQLText + " and d11_start_date >= '" + Format(frmMain.lstDatePickerFrom.Value, "YYYYMMDD") + "' "
    acSQLText = acSQLText + " and d11_end_date <= '" + Format(frmMain.lstDatePickerTo.Value, "YYYYMMDD") + "' "
    acSQLText = acSQLText + " and d11_unit_activity.agency_Code = 'ELEC' "
    acSQLText = acSQLText + " and ((d11_unit_activity.unit < '199' or d11_unit_activity.unit > '251') and d11_unit_activity.unit <> '637')"
    acSQLText = acSQLText + " order by CALL_TYPE_CD, UNIT, D11_START_DATE, D11_START_TIME "

    frmMain.sbStatusBar.Panels(1).Text = "Running query..."
    adoUnitList.Open acSQLText, adoConnection, adOpenStatic, adLockReadOnly, adCmdText

    If Not (adoUnitList.BOF And adoUnitList.EOF) Then
        iTotalRows = adoUnitList.RecordCount
          
        If iTotalRows > 0 Then
            ' Create each worksheet and the column headers.
            Set excelWorksheet = excelWorkBook.Worksheets(1)
            excelWorksheet.Name = "Priority"
            excelWorksheet.Select
            CreateWorksheetHeader excelWorksheet, excelWorkBook, excelExcelApp
            Set excelWorksheet = excelWorkBook.Worksheets(2)
            excelWorksheet.Name = "Routine"
            excelWorksheet.Select
            CreateWorksheetHeader excelWorksheet, excelWorkBook, excelExcelApp
            Set excelWorksheet = excelWorkBook.Worksheets(3)
            excelWorksheet.Select
            excelWorksheet.Name = "Street Lights-Other"
            CreateWorksheetHeader excelWorksheet, excelWorkBook, excelExcelApp
           
            'Update the progress bar
            frmMain.ProgressBar1.Max = iTotalRows + 1
            frmMain.ProgressBar1.Min = 1
            frmMain.ProgressBar1.Visible = True
            
            i = 0
            iRowPosition = 2
            iStartRow = 2
            adoUnitList.MoveFirst
            acPreviousUnitNo = adoUnitList.Fields("UNIT")
            iCrewTotal = 0
            
            Set excelWorksheet = excelWorkBook.Worksheets(1)
            Do While Not adoUnitList.EOF
                If excelWorksheet.Name <> "Routine" And adoUnitList.Fields("CALL_TYPE_CD") = "EROU" Then
                    Set excelWorksheet = excelWorkBook.Worksheets(2)
                    acPreviousUnitNo = adoUnitList.Fields("UNIT")
                    iCrewTotal = 0
                    iRowPosition = 2
                    iStartRow = 2
                ElseIf excelWorksheet.Name <> "Street Lights-Other" And adoUnitList.Fields("CALL_TYPE_CD") = "STRL" Then
                    Set excelWorksheet = excelWorkBook.Worksheets(3)
                    acPreviousUnitNo = adoUnitList.Fields("UNIT")
                    iCrewTotal = 0
                    iRowPosition = 2
                    iStartRow = 2
                End If

                acIsrNo = adoUnitList.Fields.Item("ISR_NO")
                acReportingTime = adoUnitList.Fields("FORMATED_START_TIME")
                acCreatedTime = adoUnitList.Fields("CREATION_TIME")
                acUnitNo = adoUnitList.Fields("UNIT")
                
                excelWorksheet.Cells(iRowPosition, CAD_ID_COLUMN) = acIsrNo
                excelWorksheet.Cells(iRowPosition, UNIT_COLUMN) = acUnitNo
                excelWorksheet.Cells(iRowPosition, PROBLEM_CODE_COLUMN) = adoUnitList.Fields("INIT_SERVICE_CODE")
                excelWorksheet.Cells(iRowPosition, REFER_TO_COLUMN) = adoUnitList.Fields("P_FILTER1")
                excelWorksheet.Cells(iRowPosition, CREATED_DATE_COLUMN) = acCreatedTime
                excelWorksheet.Cells(iRowPosition, REPORT_DATE_COLUMN) = acReportingTime
                excelWorksheet.Cells(iRowPosition, TOTAL_CALLS_COLUMN) = adoUnitList.Fields("TOTAL_CALLS")
                excelWorksheet.Cells(iRowPosition, TOTAL_CUSTOMERS_COLUMN) = adoUnitList.Fields("CUSTOMERS_AFFECTED")
                
                     
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
                        excelWorksheet.Cells(iRowPosition, DISPATCH_DATE_COLUMN) = acDispatchTime
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
                        excelWorksheet.Cells(iRowPosition, ACCEPTED_DATE_COLUMN) = acAcceptedTime
                    End If
                End If
                adoStatusTime.Close
                
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
                        excelWorksheet.Cells(iRowPosition, ENROUTE_DATE_COLUMN) = acEnrouteTime
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
                        excelWorksheet.Cells(iRowPosition, ARRIVED_DATE_COLUMN) = acArrivedTime
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
                        excelWorksheet.Cells(iRowPosition, CLEARED_DATE_COLUMN) = acClearedTime
                    End If
                End If
                adoStatusTime.Close
                
                '* Put the formulas for the time differences columns
                excelWorksheet.Cells(iRowPosition, CREATE_TO_ASSIGN_COLUMN) = "=+(RC[-" + Trim(Str(CREATE_TO_ASSIGN_COLUMN - DISPATCH_DATE_COLUMN)) + "]-RC[-" + Trim(Str(CREATE_TO_ASSIGN_COLUMN - CREATED_DATE_COLUMN)) + "])*1440"
                excelWorksheet.Cells(iRowPosition, ASSIGNED_TO_ARRIVED_COLUMN) = "=+(RC[-" + Trim(Str(ASSIGNED_TO_ARRIVED_COLUMN - ARRIVED_DATE_COLUMN)) + "]-RC[-" + Trim(Str(ASSIGNED_TO_ARRIVED_COLUMN - DISPATCH_DATE_COLUMN)) + "])*1440"
                excelWorksheet.Cells(iRowPosition, CREATE_TO_CLEARED_COLUMN) = "=+(RC[-" + Trim(Str(CREATE_TO_CLEARED_COLUMN - CLEARED_DATE_COLUMN)) + "]-RC[-" + Trim(Str(CREATE_TO_CLEARED_COLUMN - CREATED_DATE_COLUMN)) + "])*1440"
                excelWorksheet.Cells(iRowPosition, CREATE_TO_ARRIVED_COLUMN) = "=+(RC[-" + Trim(Str(CREATE_TO_ARRIVED_COLUMN - ARRIVED_DATE_COLUMN)) + "]-RC[-" + Trim(Str(CREATE_TO_ARRIVED_COLUMN - CREATED_DATE_COLUMN)) + "])*1440"
                excelWorksheet.Cells(iRowPosition, ENROUTE_TO_ARRIVED_COLUMN) = "=+(RC[-" + Trim(Str(ENROUTE_TO_ARRIVED_COLUMN - ARRIVED_DATE_COLUMN)) + "]-RC[-" + Trim(Str(ENROUTE_TO_ARRIVED_COLUMN - ENROUTE_DATE_COLUMN)) + "])*1440"
                excelWorksheet.Cells(iRowPosition, ENROUTE_TO_CLEARED_COLUMN) = "=+(RC[-" + Trim(Str(ENROUTE_TO_CLEARED_COLUMN - CLEARED_DATE_COLUMN)) + "]-RC[-" + Trim(Str(ENROUTE_TO_CLEARED_COLUMN - ENROUTE_DATE_COLUMN)) + "])*1440"
                excelWorksheet.Cells(iRowPosition, ARRIVED_TO_CLEARED_COLUMN) = "=+(RC[-" + Trim(Str(ARRIVED_TO_CLEARED_COLUMN - CLEARED_DATE_COLUMN)) + "]-RC[-" + Trim(Str(ARRIVED_TO_CLEARED_COLUMN - ARRIVED_DATE_COLUMN)) + "])*1440"

                'Move to the next record.
                adoUnitList.MoveNext
                iRowPosition = iRowPosition + 1
                iCrewTotal = iCrewTotal + 1
                
                '* Post the summary total for the crew.
                If Not adoUnitList.EOF Then
                    If iRowPosition > 2 And acPreviousUnitNo <> adoUnitList.Fields("UNIT") Then
                        'Write a row for the crew total
                        excelWorksheet.Cells(iRowPosition, 2) = iCrewTotal
                        'Write a row for the enroute-clear total
                        excelWorksheet.Cells(iRowPosition, ENROUTE_TO_CLEARED_COLUMN) = "=SUM(R[-" + Trim(Str(iRowPosition - iStartRow)) + "]C:R[-1]C)/60"
                        excelExcelApp.Cells(iRowPosition, ENROUTE_TO_CLEARED_COLUMN).Select
                        excelExcelApp.Selection.NumberFormat = "0.00"
                        'Write the total greater than 15 for created-assigned
                        If excelWorksheet.Name = "Priority" Then
                            excelWorksheet.Cells(iRowPosition, CREATE_TO_ASSIGN_COLUMN) = "=COUNTIF(R[-" + Trim(Str(iRowPosition - iStartRow)) + "]C:R[-1]C,"">15"")"
                            excelExcelApp.Cells(iRowPosition, CREATE_TO_ASSIGN_COLUMN).Select
                            excelExcelApp.Selection.NumberFormat = "0"
                        End If
                        'Write the total greater 120 than for created-cleared
                        If excelWorksheet.Name = "Priority" Then
                            excelWorksheet.Cells(iRowPosition, CREATE_TO_CLEARED_COLUMN) = "=COUNTIF(R[-" + Trim(Str(iRowPosition - iStartRow)) + "]C:R[-1]C,"">120"")"
                            excelExcelApp.Cells(iRowPosition, CREATE_TO_CLEARED_COLUMN).Select
                            excelExcelApp.Selection.NumberFormat = "0"
                        Else
                            excelWorksheet.Cells(iRowPosition, CREATE_TO_CLEARED_COLUMN) = "=AVERAGE(R[-" + Trim(Str(iRowPosition - iStartRow)) + "]C:R[-1]C)/60"
                            excelExcelApp.Cells(iRowPosition, CREATE_TO_CLEARED_COLUMN).Select
                            excelExcelApp.Selection.NumberFormat = "0.00"
                        End If
                        'Write the total greater than 45 for the create-arrvied
                        If excelWorksheet.Name = "Priority" Then
                            excelWorksheet.Cells(iRowPosition, CREATE_TO_ARRIVED_COLUMN) = "=COUNTIF(R[-" + Trim(Str(iRowPosition - iStartRow)) + "]C:R[-1]C,"">45"")"
                            excelExcelApp.Cells(iRowPosition, CREATE_TO_ARRIVED_COLUMN).Select
                            excelExcelApp.Selection.NumberFormat = "0"
                        End If
                        'Write the total greater than 75 for the arrived-cleared
                        Select Case excelWorksheet.Name
                            Case "Priority"
                                excelWorksheet.Cells(iRowPosition, ARRIVED_TO_CLEARED_COLUMN) = "=COUNTIF(R[-" + Trim(Str(iRowPosition - iStartRow)) + "]C:R[-1]C,"">75"")"
                            Case "Routine"
                                excelWorksheet.Cells(iRowPosition, ARRIVED_TO_CLEARED_COLUMN) = "=COUNTIF(R[-" + Trim(Str(iRowPosition - iStartRow)) + "]C:R[-1]C,"">30"")"
                            Case "Street Lights-Other"
                                excelWorksheet.Cells(iRowPosition, ARRIVED_TO_CLEARED_COLUMN) = "=COUNTIF(R[-" + Trim(Str(iRowPosition - iStartRow)) + "]C:R[-1]C,"">15"")"
                        End Select
                        excelExcelApp.Cells(iRowPosition, ARRIVED_TO_CLEARED_COLUMN).Select
                        excelExcelApp.Selection.NumberFormat = "0"
                        'Highlight in row in color for the totals
                        excelWorksheet.Rows(iRowPosition).Font.ColorIndex = 3
                        
                        'Draw border for total line
                        excelWorksheet.Select
                        excelWorksheet.Rows(iRowPosition).Select
                        excelExcelApp.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
                        excelExcelApp.Selection.Borders(xlEdgeTop).Weight = xlThin
                        excelExcelApp.Selection.Borders(xlEdgeTop).ColorIndex = xlAutomatic
                        
                        iRowPosition = iRowPosition + 1
                        iCrewTotal = 0
                        iStartRow = iRowPosition
                        acPreviousUnitNo = adoUnitList.Fields("UNIT")
                    End If
                End If
                
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

    excelWorksheet.Cells(1, TOTAL_CALLS_COLUMN) = "Total Calls"
    excelWorksheet.Columns(TOTAL_CALLS_COLUMN).Select
    excelApplication.Selection.NumberFormat = "0"
    
    excelWorksheet.Cells(1, TOTAL_CUSTOMERS_COLUMN) = "Total Customers"
    excelWorksheet.Columns(TOTAL_CUSTOMERS_COLUMN).Select
    excelApplication.Selection.NumberFormat = "0"

End Sub




