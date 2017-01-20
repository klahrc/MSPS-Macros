Attribute VB_Name = "MI_Resource_Plan2"

' Description:
'
'
' Authors:      Cesar Klahr

'
' Comment
' --------------------------------------------------------------
' Initial version
'
Const xlCalculationAutomatic As Long = -4105
Const xlCalculationManual As Long = -4135


' **************************************************************
' Module Constant Declarations Follow
' **************************************************************

' Author: Cesar Klahr
' Mackenzie Investments - x5249
' Description: v0.9 Supports Late Binding to avoid issues with missing References
'
' Date: 6-Nov-2014

' Leveled Resource Plan
' MSP Resource Plan
' CCL Spreadsheet
' Resource Variance


' Show dialog to enter what the user wants to see (relative alloc, absolute alloc (this is what MSP is showing: dCellValue / PJ.Resources(r).MaxUnits, or in hours)

Const FISCAL_MONTH_ROW = 2
Const WEEKS_SEQ_ROW = 3
Const LEV_RES_PLAN_START_ROW = 5        ' Start row for the Levelled Res Plan table in the spreadsheet

Const RES_SUMMARY_START_COL = 1
Const FISCAL_MONTH_SCHED1_START_COL = 13
Const DET_SCHED1_START_COL = 26

Const SEPARATOR = 5        ' Separation among tables in the layout (in rows)
Const FIRST_COL_ALLOCATED = 13        ' First column in Resource Tracking (CCL) spreadsheet containing data

Const LEV_RES = 1
Const MSP_RES = 2
Const CCL_RES = 3
Const VAR_RES = 4

'//////////////// EARLY BINDING DECLARATIONS////////////////
'Dim m_xlApp As Excel.Application
'Dim m_wbFC As Workbook                                                                   ' Fiscal Calendar Workbook
'Dim m_wsFC As Worksheet                                                                  ' Fiscal Calendar sheet in m_wbFC

'Dim m_wbResPlan As Workbook
'Dim m_wsResPlan As Excel.Worksheet

'Dim m_wbResTracking As Workbook
'Dim m_wsResAllocated As Worksheet
'////////////////////////////////////////////////////////////

'//////////////// LATE BINDING DECLARATIONS////////////////
Dim m_xlApp As Object
Dim m_wbFC As Object        ' Fiscal Calendar Workbook
Dim m_wsFC As Object        ' Fiscal Calendar sheet in m_wbFC

Dim m_wbResPlan As Object
Dim m_wsResPlan As Object

Dim m_wbResTracking As Object
Dim m_wsResAllocated As Object
'////////////////////////////////////////////////////////////

Dim m_sProjStartDate As String        ' Will default to Project start Date
Dim m_sPrjEndDate As String        ' Will default to Project End Date

Dim m_nNumPeriods As Integer        ' Total number of project periods
Dim m_nNumResources As Integer        ' Total number of project resources

Dim MSP_Res_Plan_Start_Row As Integer        ' Start row for the MSP_Res_Plan table in the spreadsheet
Dim CCL_Spreadsheet_Start_Row As Integer        ' Start row for the CCL table in the spreadsheet
Dim Res_Variance_Start_Row As Integer        ' Start row for the Res_Variance table in the spreadsheet

Dim m_sFiscalMonthFilePath As String
Dim m_sFiscalFileName As String
Dim m_sFiscalCalendarTabName As String

Dim m_sResTrackingFilePath As String
Dim m_sResSpreadsheetFileName As String

Dim m_sProjectCode As String
Dim m_sProjectName As String

Dim m_nFirstDetScheduleWeekInFiscal As Long
Dim m_nFirstFiscalWeekNumInDetSchedule As Integer

Dim m_nLastFiscalWeekNum As Long
Dim m_nPeriodsToCopy As Long        ' Holds how many periods in the fiscal year falls within the project schedule

Dim m_sFiscalYear As String

Dim m_nDistance2NextFriday As Integer

Dim m_sCCResAllocFolder As String
Dim m_sFiscalCalFolder As String

Dim m_list As VbaList

Option Explicit
Option Base 1


Sub LoadSelectedTasks(pj1 As Project)
    Dim t As Task


    If ActiveSelection.Tasks.Count > 0 Then

        If ActiveSelection.Tasks.Count = 1 Then
            '''' IF ONLY ONE TASKS IS SELECT ASK IF THAT'S WHAT THE USER WANTS OR SHOULD I SELECT THE WHOLE PROJECT????????????????????????


        Else

            For Each t In ActiveSelection.Tasks
                m_list.Add t.ID
            Next
        End If

    Else

        For Each t In pj1.Tasks
            m_list.Add t.ID
        Next

    End If



End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Description:
'
'
' Authors:      Cesar Klahr

'
' Comment:      Copy Resource allocation from Bharat's spreadsheet
' --------------------------------------------------------------
' Initial version
'

'//////////////// EARLY BINDING DECLARATIONS////////////////
'Private Sub CopyCCLAllocations(ws1 As Worksheet)
'    Dim rngCell As Range
'    Dim rngDataRange As Range
'////////////////////////////////////////////////////////////

'//////////////// LATE BINDING DECLARATIONS////////////////
Private Sub CopyCCLAllocations(ws1 As Object)
    Dim rngCell As Object
    Dim rngDataRange As Object
    '////////////////////////////////////////////////////////////

    Dim i As Integer
    Dim lFirstRowFound As Long


    ' Open Bharat's workbook

    If m_wsResAllocated.AutoFilterMode = True Then m_wsResAllocated.AutoFilterMode = False


    ' ***Ensure NO filters are applied!
    ' ASSUMPTION: Resource Name is on Column 'G'
    Set rngDataRange = m_wsResAllocated.Columns("G")

    For i = 1 To m_nNumResources        ' For all resources in my copy of Resource tracking

        ' Look for my resource in Bharat's spreadsheet
        Set rngCell = rngDataRange.Find(What:=CStr(ws1.cells(i + CCL_Spreadsheet_Start_Row + 1, 1)), LookIn:=xlValuesLB, LookAt:=xlWholeLB, SearchOrder:=xlByRowsLB, _
                                        SearchDirection:=xlNextLB, MatchCase:=False, SearchFormat:=False)

        ' If can't find it, try again adding '*' to the name (it's a contractor)
        If (rngCell Is Nothing) Then
            Set rngCell = rngDataRange.Find(What:=CStr(ws1.cells(i + CCL_Spreadsheet_Start_Row + 1, 1)) + "*", LookIn:=xlValuesLB, LookAt:=xlWholeLB, SearchOrder:=xlByRowsLB, _
                                            SearchDirection:=xlNextLB, MatchCase:=False, SearchFormat:=False)
        End If

        If Not (rngCell Is Nothing) Then        ' Found it!
            lFirstRowFound = rngCell.row

            Do Until False        ' Infinite loop until....found my project code or started over!
                ' Check if the allocation is for my project
                With m_wsResAllocated
                    If .cells(rngCell.row, 1) = m_sProjectCode Then

                        ' Copy allocations to my spreadsheet!

                        ' + m_nFirstFiscalWeekNumInDetSchedule - 1??????
                        .Range(.cells(rngCell.row, FIRST_COL_ALLOCATED + m_nFirstFiscalWeekNumInDetSchedule - 1), _
                               .cells(rngCell.row, FIRST_COL_ALLOCATED + m_nFirstFiscalWeekNumInDetSchedule - 1 + m_nLastFiscalWeekNum - m_nFirstFiscalWeekNumInDetSchedule)).Copy
                        ws1.cells(i + CCL_Spreadsheet_Start_Row + 1, DET_SCHED1_START_COL + m_nFirstDetScheduleWeekInFiscal - 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormatsLB


                        ws1.cells(i + CCL_Spreadsheet_Start_Row + 1, 2) = "ASSIGNED"        ' Update resource status
                        ws1.cells(i + CCL_Spreadsheet_Start_Row + 1, 2).Font.Color = vbBlack


                        ws1.cells(i + LEV_RES_PLAN_START_ROW + 1, 2) = rngCell.Offset(, 1)
                        ws1.cells(i + LEV_RES_PLAN_START_ROW + 1 + m_nNumResources + SEPARATOR, 2) = rngCell.Offset(, 1)

                        '''''''''''ws1.Cells(i, 2) = 1

                        Exit Do        ' There shouldn't be more than 1 row for this resource for my project in Bharat's spreadsheet!

                    End If
                End With

                'It's not my project...so let's move to the next row found...
                Set rngCell = rngDataRange.FindNext(rngCell)

                If (rngCell.row <= lFirstRowFound) Then
                    ws1.cells(i + CCL_Spreadsheet_Start_Row + 1, 2) = "PENDING"        ' Update resource status
                    ws1.cells(i + CCL_Spreadsheet_Start_Row + 1, 2).Font.Color = vbMagenta

                    ws1.cells(i + LEV_RES_PLAN_START_ROW + 1, 2) = rngCell.Offset(, 1)
                    ws1.cells(i + LEV_RES_PLAN_START_ROW + 1 + m_nNumResources + SEPARATOR, 2) = rngCell.Offset(, 1)
                    Exit Do        '...until it starts looping...
                End If

            Loop

        Else        ' Resource not found in Bharat's spreadsheet!
            ws1.cells(i + CCL_Spreadsheet_Start_Row + 1, 2) = "NOT FOUND"        ' Update resource status
            ws1.cells(i + CCL_Spreadsheet_Start_Row + 1, 2).Font.Color = vbRed

            ws1.cells(i + LEV_RES_PLAN_START_ROW + 1, 2) = "N/A"
            ws1.cells(i + LEV_RES_PLAN_START_ROW + 1 + m_nNumResources + SEPARATOR, 2) = "N/A"
        End If
    Next        ' Next resource...

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Description:
'
'
' Authors:      Cesar Klahr

'
' Comment:      Clean Resource allocation from Bharat's spreadsheet
' --------------------------------------------------------------
' Initial version
'
'//////////////// EARLY BINDING DECLARATIONS////////////////
'Private Sub CleanUPCCLAllocations(ws1 As Worksheet)
'////////////////////////////////////////////////////////////

'//////////////// LATE BINDING DECLARATIONS////////////////
Private Sub CleanUPCCLAllocations(ws1 As Object)
    '////////////////////////////////////////////////////////////

    Dim i As Integer
    Dim j As Integer

    For i = CCL_Spreadsheet_Start_Row + 2 To CCL_Spreadsheet_Start_Row + 2 + m_nNumResources - 1
        For j = DET_SCHED1_START_COL + m_nFirstDetScheduleWeekInFiscal - 1 To DET_SCHED1_START_COL + m_nFirstDetScheduleWeekInFiscal - 1 + m_nPeriodsToCopy - 1
            If Not (IsNumber(ws1.cells(i, j))) Then
                ws1.cells(i, j) = ""
            End If
        Next
    Next

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Description:
'
'
' Authors:      Cesar Klahr

'
' Comment
' --------------------------------------------------------------
' Initial version
'
'//////////////// EARLY BINDING DECLARATIONS////////////////
' Private Sub Draw_tables(pj1 As Project, ws1 As Excel.Worksheet, nStartRow As Integer, TSVWork1 As TimeScaleValues, nTableType As Integer, tsu As PjTimescaleUnit)
'////////////////////////////////////////////////////////////

'//////////////// LATE BINDING DECLARATIONS////////////////
Private Sub Draw_tables(pj1 As Project, ws1 As Object, nStartRow As Integer, TSVWork1 As TimeScaleValues, nTableType As Integer, tsu As PjTimescaleUnit)
    '////////////////////////////////////////////////////////////

    Dim t As Long
    Dim r As Long
    Dim d As Integer
    Dim dCellValue As Double
    Dim dRoundUp As Double
    Dim dRoundDown As Double
    Dim dTotResAlloc As Double
    Dim j As Integer
    Dim nCurRow As Integer
    Dim nCurCol As Integer

    Dim sOldMonth As String
    Dim nWeeksInMonth As Integer
    Dim nWeeks As Integer
    Dim bMonthChanged As Boolean
    Dim avMonths As Variant
    Dim nMonths As Integer
    Dim sFormulaSumIf As String
    Dim sRate As String
    Dim asName() As String
    Dim bReverse As Boolean
    Dim lFrom As Long

    '//////////////// EARLY BINDING DECLARATIONS////////////////
    'Dim rng As Range
    'Dim rngCell As Range
    'Dim rngDataRange As Range
    '////////////////////////////////////////////////////////////

    '//////////////// LATE BINDING DECLARATIONS////////////////
    Dim rng As Object
    Dim rngCell As Object
    Dim rngDataRange As Object
    '////////////////////////////////////////////////////////////

    avMonths = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

    ' Choose work unit divisor depending on the Tools | Options | Schedule | Work.
    ' Work is stored in minutes in MS Project.
    Select Case pj1.DefaultWorkUnits
        Case pjMinute
            d = 1
        Case pjHour
            d = 60
        Case pjDay
            d = pj1.HoursPerDay * 60
        Case pjWeek
            d = pj1.HoursPerWeek * 60
        Case pjMonthUnit
            d = pj1.DaysPerMonth * pj1.HoursPerDay * 60
        Case Else
            d = 1
    End Select







    With ws1

        ' Draw Grid for Resource summary info table
        With .Range(.cells(nStartRow + 1, RES_SUMMARY_START_COL), _
                    .cells(nStartRow + 1 + m_nNumResources, RES_SUMMARY_START_COL + 10))
            .Borders.LineStyle = xlContinuousLB
            '''''.Borders.Weight = xlThinLB
        End With

        ' Setting defaults for Resource summary Grid headers
        .Range(.cells(nStartRow + 1, RES_SUMMARY_START_COL), .cells(nStartRow + 1, RES_SUMMARY_START_COL + 10)).Interior.Color = RGB(15, 37, 63)
        .Range(.cells(nStartRow + 1, RES_SUMMARY_START_COL), .cells(nStartRow + 1, RES_SUMMARY_START_COL + 10)).Font.Color = vbWhite
        .Range(.cells(nStartRow + 1, RES_SUMMARY_START_COL), .cells(nStartRow + 1, RES_SUMMARY_START_COL + 10)).Font.Bold = True


        ' Paint the Resource summary Grid with Gray
        .Range(.cells(nStartRow + 2, RES_SUMMARY_START_COL), _
               .cells(nStartRow + 2 + m_nNumResources, RES_SUMMARY_START_COL + 10)).Interior.Color = RGB(219, 229, 241)


        Select Case nTableType

            Case LEV_RES
                .cells(nStartRow, RES_SUMMARY_START_COL) = "Levelled Resource Plan"
            Case MSP_RES
                .cells(nStartRow, RES_SUMMARY_START_COL) = "MSP Resource Plan"
            Case CCL_RES
                .cells(nStartRow, RES_SUMMARY_START_COL) = "CCL Resource Plan"
            Case VAR_RES
                .cells(nStartRow, RES_SUMMARY_START_COL) = "Variance Resource Plan"
        End Select
        .cells(nStartRow, RES_SUMMARY_START_COL).Font.Size = 10
        .cells(nStartRow, RES_SUMMARY_START_COL).Font.Bold = True
        .cells(nStartRow, RES_SUMMARY_START_COL).Font.Color = vbWhite
        .cells(nStartRow, RES_SUMMARY_START_COL).Interior.Color = RGB(83, 142, 213)
        .cells(nStartRow, RES_SUMMARY_START_COL).VerticalAlignment = xlTop
        .Rows(nStartRow).RowHeight = 20

        .cells(nStartRow + 1, RES_SUMMARY_START_COL) = "Resource Name"
        .cells(nStartRow + 1, RES_SUMMARY_START_COL).ColumnWidth = 20

        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 1) = "Status"
        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 1).ColumnWidth = 10

        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 2) = "Company"
        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 2).AddComment ("Company inferred from email address field in MSP")
        '      .cells(nStartRow + 1, RES_SUMMARY_START_COL + 2).Comment.Shape.TextFrame.Characters.Font.Bold = False
        '      .cells(nStartRow + 1, RES_SUMMARY_START_COL + 2).Comment.Shape.TextFrame.Characters.Font.Name = "Calibri"
        '      .cells(nStartRow + 1, RES_SUMMARY_START_COL + 2).Comment.Shape.TextFrame.Characters.Font.Size = 8
        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 2).ColumnWidth = 10

        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 3) = "Group"
        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 3).ColumnWidth = 20

        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 4) = "Rate"
        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 4).AddComment ("Rate derived from StandardRate field in MSP")
        '.cells(nStartRow + 1, RES_SUMMARY_START_COL + 4).Comment.Shape.TextFrame.Characters.Font.Bold = False
        '.cells(nStartRow + 1, RES_SUMMARY_START_COL + 4).Comment.Shape.TextFrame.Characters.Font.Name = "Calibri"
        '.cells(nStartRow + 1, RES_SUMMARY_START_COL + 4).Comment.Shape.TextFrame.Characters.Font.Size = 8
        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 4).ColumnWidth = 10
        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 4).HorizontalAlignment = xlRight


        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 5) = "In Year FTE"
        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 5).AddComment ("Average FTE. Should match CCL spreadsheet")
        '.cells(nStartRow + 1, RES_SUMMARY_START_COL + 5).Comment.Shape.TextFrame.Characters.Font.Bold = False
        '.cells(nStartRow + 1, RES_SUMMARY_START_COL + 5).Comment.Shape.TextFrame.Characters.Font.Name = "Calibri"
        '.cells(nStartRow + 1, RES_SUMMARY_START_COL + 5).Comment.Shape.TextFrame.Characters.Font.Size = 8
        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 5).ColumnWidth = 10
        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 5).HorizontalAlignment = xlRight

        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 6) = "In Year Hrs"
        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 6).ColumnWidth = 10
        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 6).HorizontalAlignment = xlRight

        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 7) = "In Year Cost"
        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 7).ColumnWidth = 12
        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 7).HorizontalAlignment = xlRight

        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 8) = "Total FTE"
        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 8).AddComment ("Average FTE. Should match CCL spreadsheet")
        '.cells(nStartRow + 1, RES_SUMMARY_START_COL + 8).Comment.Shape.TextFrame.Characters.Font.Bold = False
        '.cells(nStartRow + 1, RES_SUMMARY_START_COL + 8).Comment.Shape.TextFrame.Characters.Font.Name = "Calibri"
        '.cells(nStartRow + 1, RES_SUMMARY_START_COL + 8).Comment.Shape.TextFrame.Characters.Font.Size = 8
        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 8).ColumnWidth = 10
        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 8).HorizontalAlignment = xlRight

        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 9) = "Total Hrs"
        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 9).ColumnWidth = 10
        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 9).HorizontalAlignment = xlRight

        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 10) = "Total Cost"
        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 10).ColumnWidth = 12
        .cells(nStartRow + 1, RES_SUMMARY_START_COL + 10).HorizontalAlignment = xlRight


        ' Draw Grid for Fiscal Month schedule (12 Months)
        With .Range(.cells(nStartRow, FISCAL_MONTH_SCHED1_START_COL), _
                    .cells(nStartRow + m_nNumResources + 1, FISCAL_MONTH_SCHED1_START_COL + 11))
            .Borders.LineStyle = xlContinuousLB
            '''            .Borders.Weight = xlThinLB
            .ColumnWidth = 6
        End With


        ' Add Year in header of Fiscal Month schedule and merge cells
        .cells(nStartRow, FISCAL_MONTH_SCHED1_START_COL) = "FISCAL CALENDAR - " + m_sFiscalYear
        .Range(.cells(nStartRow, FISCAL_MONTH_SCHED1_START_COL), _
               .cells(nStartRow, FISCAL_MONTH_SCHED1_START_COL + 11)).Merge
        .Range(.cells(nStartRow, FISCAL_MONTH_SCHED1_START_COL), _
               .cells(nStartRow, FISCAL_MONTH_SCHED1_START_COL + 11)).HorizontalAlignment = xlCenter
        .Range(.cells(nStartRow, FISCAL_MONTH_SCHED1_START_COL), _
               .cells(nStartRow, FISCAL_MONTH_SCHED1_START_COL + 11)).Font.Size = 10
        .Range(.cells(nStartRow, FISCAL_MONTH_SCHED1_START_COL), _
               .cells(nStartRow, FISCAL_MONTH_SCHED1_START_COL + 11)).VerticalAlignment = xlTop
        .Range(.cells(nStartRow, FISCAL_MONTH_SCHED1_START_COL), _
               .cells(nStartRow, FISCAL_MONTH_SCHED1_START_COL + 11)).Interior.Color = RGB(149, 55, 53)
        .Range(.cells(nStartRow, FISCAL_MONTH_SCHED1_START_COL), _
               .cells(nStartRow, FISCAL_MONTH_SCHED1_START_COL + 11)).Font.Color = vbWhite
        .Range(.cells(nStartRow, FISCAL_MONTH_SCHED1_START_COL), _
               .cells(nStartRow, FISCAL_MONTH_SCHED1_START_COL + 11)).Font.Bold = True


        ' Add Months in header of Fiscal Month schedule
        For nMonths = 1 To 12
            .cells(nStartRow + 1, FISCAL_MONTH_SCHED1_START_COL + nMonths - 1) = avMonths(nMonths)
        Next

        .Range(.cells(nStartRow + 1, FISCAL_MONTH_SCHED1_START_COL), _
               .cells(nStartRow + 1, FISCAL_MONTH_SCHED1_START_COL + 11)).HorizontalAlignment = xlCenter
        .Range(.cells(nStartRow + 1, FISCAL_MONTH_SCHED1_START_COL), _
               .cells(nStartRow + 1, FISCAL_MONTH_SCHED1_START_COL + 11)).Interior.Color = RGB(219, 229, 241)        ' Gray




        ' Add formula for Fiscal Month cells
        '=SUMIF(R2C26:R2C56,"="&R6C,RC26:RC55)/4
        sFormulaSumIf = "=SUMIF(R" + CStr(FISCAL_MONTH_ROW) + "C" + CStr(DET_SCHED1_START_COL) + ":" + "R" + CStr(FISCAL_MONTH_ROW) + "C" + CStr(DET_SCHED1_START_COL + m_nNumPeriods) + _
                        "," + """=""" + "&" + "R" + CStr(nStartRow + 1) + "C,RC" + CStr(DET_SCHED1_START_COL) + ":RC" + CStr(DET_SCHED1_START_COL + m_nNumPeriods) + ")/4"
        .Range(.cells(nStartRow + 2, FISCAL_MONTH_SCHED1_START_COL), _
               .cells(nStartRow + 2 + m_nNumResources, FISCAL_MONTH_SCHED1_START_COL + 11)) = sFormulaSumIf
        .Range(.cells(nStartRow + 2, FISCAL_MONTH_SCHED1_START_COL), _
               .cells(nStartRow + 2 + m_nNumResources, FISCAL_MONTH_SCHED1_START_COL + 11)).NumberFormat = "#,##0.00"

        ' Paint summary row in Fiscal Month table
        .Range(.cells(nStartRow + 2 + m_nNumResources, FISCAL_MONTH_SCHED1_START_COL), _
               .cells(nStartRow + 2 + m_nNumResources, FISCAL_MONTH_SCHED1_START_COL + 11)).Interior.Color = RGB(15, 37, 63)
        .Range(.cells(nStartRow + 2 + m_nNumResources, FISCAL_MONTH_SCHED1_START_COL), _
               .cells(nStartRow + 2 + m_nNumResources, FISCAL_MONTH_SCHED1_START_COL + 11)).Font.Color = vbWhite

        bReverse = (nTableType = VAR_RES)        ' I want to highlith the opposite in the Resource variancetable
        ' Conditional formatting for the Fiscal Month Grid
        Call AddConditionalFormatting(ws1, nStartRow + 2, FISCAL_MONTH_SCHED1_START_COL, nStartRow + 2 + m_nNumResources - 1, FISCAL_MONTH_SCHED1_START_COL + 11, bReverse)


        ' Print dates for detailed schedule
        ' Figure out Fiscal Month for each week in the Detailed schedule
        '=VLOOKUP(Z3,'C:\Users\cklahr\Documents\[Fiscal Calendar.xlsx]Fiscal Calendar'!$J$4:$M$56, 4, FALSE)


        ' Draw Grid for Detailed Schedule
        With .Range(.cells(nStartRow, DET_SCHED1_START_COL), .cells(nStartRow + m_nNumResources + 1, DET_SCHED1_START_COL + m_nNumPeriods - 1))
            .Borders.LineStyle = xlContinuousLB
            .Borders.Weight = xlThinLB
            .ColumnWidth = 4
        End With

        ' Paint headers with Gray
        .Range(.cells(nStartRow, DET_SCHED1_START_COL), _
               .cells(nStartRow + 1, DET_SCHED1_START_COL + m_nNumPeriods - 1)).Interior.Color = RGB(219, 229, 241)

        ' Add Dates and Months in the headers
        nCurRow = nStartRow
        nCurCol = DET_SCHED1_START_COL
        nWeeksInMonth = 0
        bMonthChanged = False


        If m_nNumPeriods > 0 Then sOldMonth = Format(TSVWork1(1).StartDate + m_nDistance2NextFriday, "mmm-yy")        ' sOldMonth is the month of the first period

        For t = 1 To m_nNumPeriods

            ' Put the month in the first row of the header
            bMonthChanged = sOldMonth <> Format(TSVWork1(t).StartDate + m_nDistance2NextFriday, "mmm-yy")

            If bMonthChanged Then        ' Changed Month
                ws1.cells(nCurRow, nCurCol - nWeeksInMonth) = "01-" + sOldMonth        ' Populate last month in the header
                ws1.cells(nCurRow, nCurCol - nWeeksInMonth).NumberFormat = "mmm-yy"
                .Range(.cells(nCurRow, nCurCol - nWeeksInMonth), .cells(nCurRow, nCurCol - 1)).Merge
                .Range(.cells(nCurRow, nCurCol - nWeeksInMonth), .cells(nCurRow, nCurCol - 1)).HorizontalAlignment = xlCenter

                ' Make borders thicker for the entire month
                With .Range(.cells(nCurRow, nCurCol - nWeeksInMonth), .cells(nCurRow + m_nNumResources + 1, nCurCol - 1))
                    .Borders(xlEdgeTopLB).Weight = xlThickLB
                    .Borders(xlEdgeLeftLB).Weight = xlThickLB
                    .Borders(xlEdgeRightLB).Weight = xlThickLB
                    .Borders(xlEdgeBottomLB).Weight = xlThickLB
                End With
                nWeeksInMonth = 0        ' Reset numbers of weeks for next month

                sOldMonth = Format(TSVWork1(t).StartDate + m_nDistance2NextFriday, "mmm-yy")        ' Store new month in the oldMonth variable
            End If

            ' Special treatment for the last period
            If t = m_nNumPeriods Then
                ' In this case, I put current month in the header instead of old month
                ws1.cells(nCurRow, nCurCol - nWeeksInMonth) = "01-" + Format(TSVWork1(t).StartDate + m_nDistance2NextFriday, "mmm-yy")
                ws1.cells(nCurRow, nCurCol - nWeeksInMonth).NumberFormat = "mmm-yy"
                .Range(.cells(nCurRow, nCurCol - nWeeksInMonth), .cells(nCurRow, nCurCol)).Merge
                .Range(.cells(nCurRow, nCurCol - nWeeksInMonth), .cells(nCurRow, nCurCol)).HorizontalAlignment = xlCenter

                ' Make borders thicker for the partial month
                With .Range(.cells(nCurRow, nCurCol - nWeeksInMonth), .cells(nCurRow + m_nNumResources + 1, nCurCol))
                    .Borders(xlEdgeTopLB).Weight = xlThickLB
                    .Borders(xlEdgeLeftLB).Weight = xlThickLB
                    .Borders(xlEdgeRightLB).Weight = xlThickLB
                    .Borders(xlEdgeBottomLB).Weight = xlThickLB
                End With
            End If

            ' Put the day of the month in the second row of the header
            ws1.cells(nCurRow + 1, nCurCol) = TSVWork1(t).StartDate + m_nDistance2NextFriday
            ws1.cells(nCurRow + 1, nCurCol).NumberFormat = "dd"

            ' Increment number of weeks within the month
            nWeeksInMonth = nWeeksInMonth + 1
            nCurCol = nCurCol + 1
        Next t

        ' ************** THIS SHOULD BE DONE JUST ONCE !!!! ***********************
        ' First week in the Fiscal Calendar falling into the Detailed schedule (m_nFirstFiscalWeekNumInDetSchedule)
        ' First week in the Detailed schedule falling into the Fiscal Calendar (m_nFirstDetScheduleWeekInFiscal)
        Set rngDataRange = m_wsFC.Range("L:L")

        Dim bFoundDateInFiscalCalendar As Boolean

        bFoundDateInFiscalCalendar = False
        For nWeeks = 1 To m_nNumPeriods
            Set rngCell = rngDataRange.Find(What:=.cells(LEV_RES_PLAN_START_ROW + 1, DET_SCHED1_START_COL + nWeeks - 1), LookIn:=xlValuesLB, LookAt:=xlWholeLB, SearchOrder:=xlByRowsLB, SearchDirection:=xlNextLB, _
                                            MatchCase:=False, SearchFormat:=False)

            If Not (rngCell Is Nothing) Then
                ' Found it
                m_nFirstFiscalWeekNumInDetSchedule = rngCell.Offset(, -2)
                m_nFirstDetScheduleWeekInFiscal = nWeeks
                bFoundDateInFiscalCalendar = True
                Exit For
            End If
        Next

        If Not bFoundDateInFiscalCalendar Then
            '''''Set pj = Nothing       ' Review later how to free up memory for this!!!
            '''''Set TSVWork = Nothing  ' Review later how to free up memory for this!!!
            m_wbFC.Close savechanges:=False
            m_wbResTracking.Close savechanges:=False
            m_wbResPlan.Close savechanges:=False
            m_xlApp.Quit

            Set m_xlApp = Nothing
            Set m_wbResPlan = Nothing
            Set m_wsResPlan = Nothing
            MsgBox ("Can't find project dates in the Fiscal Calendar provided. Please check and try again later")
            End
        End If

        ' This is the magic formula!!!
        m_nPeriodsToCopy = m_xlApp.WorksheetFunction.Min(m_nNumPeriods - m_nFirstDetScheduleWeekInFiscal + 1, m_nLastFiscalWeekNum - m_nFirstFiscalWeekNumInDetSchedule + 1)


        For nWeeks = 1 To m_nPeriodsToCopy
            .cells(FISCAL_MONTH_ROW, DET_SCHED1_START_COL + nWeeks - 1 + m_nFirstDetScheduleWeekInFiscal - 1) = "=VLOOKUP(R[1]C,'" + m_sFiscalMonthFilePath + "[" + m_sFiscalFileName + "]" + m_sFiscalCalendarTabName + "'!C10:C13, 4, FALSE)"
            .cells(FISCAL_MONTH_ROW, DET_SCHED1_START_COL + nWeeks - 1 + m_nFirstDetScheduleWeekInFiscal - 1).Font.Color = vbRed
        Next

        ' Add Weeks in sequence in row WEEKS_SEQ_ROW
        For nWeeks = 1 To m_nPeriodsToCopy
            .cells(WEEKS_SEQ_ROW, DET_SCHED1_START_COL + nWeeks - 1 + m_nFirstDetScheduleWeekInFiscal - 1) = m_nFirstFiscalWeekNumInDetSchedule + nWeeks - 1
        Next

        '***************************************************************************

        ' Now paint even and odd months with different colors
        For t = 1 To m_nNumPeriods
            If (.cells(nStartRow + 1, t + DET_SCHED1_START_COL - 1) < DateValue(Now())) Then
                .Range(.cells(nStartRow + 2, t + DET_SCHED1_START_COL - 1), .cells(nStartRow + m_nNumResources + 1, t + DET_SCHED1_START_COL - 1)).Interior.Color = RGB(219, 229, 241)
            ElseIf (Month(.cells(nStartRow + 1, t + DET_SCHED1_START_COL - 1)) Mod 2) = 0 Then
                .Range(.cells(nStartRow + 2, t + DET_SCHED1_START_COL - 1), .cells(nStartRow + m_nNumResources + 1, t + DET_SCHED1_START_COL - 1)).Interior.Color = RGB(153, 255, 153)        '122, 197, 205
            Else
                .Range(.cells(nStartRow + 2, t + DET_SCHED1_START_COL - 1), .cells(nStartRow + m_nNumResources + 1, t + DET_SCHED1_START_COL - 1)).Interior.Color = RGB(212, 226, 184)        ' (187, 209, 143)
            End If
        Next
    End With


    Dim tempAsgnEffort() As Double
    ReDim tempAsgnEffort(m_nNumPeriods)


    Dim Asgn As Assignment
    Dim TSVPlannedWork As TimeScaleValues

    Dim lShouldBeAdded As Boolean

    ' Start populating Resource info in Summary and Detailed tables
    For r = 1 To pj1.ResourceCount

        For t = 1 To m_nNumPeriods
            tempAsgnEffort(t) = 0
        Next

        lShouldBeAdded = False

        For Each Asgn In pj1.Resources(r).Assignments

            ' If the assignment is for a task in the selection, THEN ADD THE RESOURCE AND TIMEPHASED DATA!!!
            If m_list.Exists(Asgn.Task.ID) Then
                lShouldBeAdded = True


                Set TSVPlannedWork = Asgn.TimeScaleData(m_sProjStartDate, m_sPrjEndDate, _
                                                        Type:=pjAssignmentTimescaledWork, TimescaleUnit:=tsu)

                For t = 1 To m_nNumPeriods

                    If Not TSVPlannedWork(t).Value = "" And Not TSVPlannedWork(t).Value = 0 Then

                        tempAsgnEffort(t) = tempAsgnEffort(t) + TSVPlannedWork(t).Value / d / 37.5

                    End If
                Next
            End If

        Next Asgn


        If lShouldBeAdded Then




            nCurCol = DET_SCHED1_START_COL

            '''''''''''''''''''''Set TSVWork1 = pj1.Resources(r).TimeScaleData(m_sProjStartDate, m_sPrjEndDate, _
             Type:=pjResourceTimescaledWork, TimescaleUnit:=tsu)

            ' Populate info i Summary Table
            asName() = Split(pj1.Resources(r).Name, ";")
            If UBound(asName) > 0 Then
                ws1.cells(nCurRow + 2, 1) = Trim(asName(1)) + " " + Trim(asName(0))
            Else
                ws1.cells(nCurRow + 2, 1) = pj1.Resources(r).Name
            End If


            If InStr(1, UCase(pj1.Resources(r).EMailAddress), "MACKENZIE") > 0 Then
                ws1.cells(nCurRow + 2, 3) = "MFC"
            ElseIf InStr(1, UCase(pj1.Resources(r).EMailAddress), "INVESTORS") > 0 Then
                ws1.cells(nCurRow + 2, 3) = "IG"
            ElseIf InStr(1, UCase(pj1.Resources(r).EMailAddress), "LONDON") > 0 Then
                ws1.cells(nCurRow + 2, 3) = "London Life"
            ElseIf InStr(1, UCase(pj1.Resources(r).EMailAddress), "GWL") > 0 Then
                ws1.cells(nCurRow + 2, 3) = "GWL"
            ElseIf InStr(1, UCase(pj1.Resources(r).EMailAddress), "CANADA") > 0 Then
                ws1.cells(nCurRow + 2, 3) = "Canada Life"
            End If

            ws1.cells(nCurRow + 2, 4) = pj1.Resources(r).Group

            sRate = pj1.Resources(r).StandardRate
            sRate = Replace(sRate, "$", "")
            sRate = Replace(sRate, "/hr", "")

            If sRate <> "" Then ws1.cells(nCurRow + 2, 5) = Format(sRate, "$#,##0.00")

            ' *********************** Apply to the whole table!!!!!!!!!!!!!!!!!! *******************************
            lFrom = DET_SCHED1_START_COL + m_nFirstDetScheduleWeekInFiscal - 1
            ws1.cells(nCurRow + 2, 6) = "=IF(SUM(RC[" + CStr(lFrom - 6) + "]:RC[" + CStr(lFrom + m_nPeriodsToCopy - 6 - 1) + _
                                        "])>0, SUM(RC[" + CStr(lFrom - 6) + "]:RC[" + CStr(lFrom + m_nPeriodsToCopy - 6 - 1) + _
                                        "])/COUNTIF(RC[" + CStr(lFrom - 6) + "]:RC[" + CStr(lFrom + m_nPeriodsToCopy - 6 - 1) + "]," + _
                                        """>0""" + "),0)"

            ws1.cells(nCurRow + 2, 6).NumberFormat = "#,##0.00"

            ws1.cells(nCurRow + 2, 7) = "=SUM(RC[" + CStr(lFrom - 7) + "]:RC[" + CStr(lFrom + m_nPeriodsToCopy - 7 - 1) + "])*37.5"
            ws1.cells(nCurRow + 2, 7).NumberFormat = "#,##0.00"

            ws1.cells(nCurRow + 2, 8) = "=RC[-1]*RC[-3]"
            ws1.cells(nCurRow + 2, 8).NumberFormat = "$#,##0.00"

            ws1.cells(nCurRow + 2, 9) = "=IF(SUM(RC[" + CStr(DET_SCHED1_START_COL - 9) + "]:RC[" + CStr(DET_SCHED1_START_COL + m_nNumPeriods - 9 - 1) + "])>0, SUM(RC[" + _
                                        CStr(DET_SCHED1_START_COL - 9) + "]:RC[" + CStr(DET_SCHED1_START_COL + m_nNumPeriods - 9 - 1) + "])/COUNTIF(RC[" + _
                                        CStr(DET_SCHED1_START_COL - 9) + "]:RC[" + CStr(DET_SCHED1_START_COL + m_nNumPeriods - 9 - 1) + "]," + """>0""" + "),0)"

            ws1.cells(nCurRow + 2, 9).NumberFormat = "#,##0.00"

            ws1.cells(nCurRow + 2, 10) = "=SUM(RC[" + CStr(DET_SCHED1_START_COL - 10) + "]:RC[" + CStr(DET_SCHED1_START_COL + m_nNumPeriods - 10 - 1) + "])*37.5"
            ws1.cells(nCurRow + 2, 10).NumberFormat = "#,##0.00"

            ws1.cells(nCurRow + 2, 11) = "=RC[-1]*RC[-6]"
            ws1.cells(nCurRow + 2, 11).NumberFormat = "$#,##0.00"
            ' ****************************************************************************************************

            ' Now populate info in Detailed Schedule Table




            dTotResAlloc = 0
            For t = 1 To m_nNumPeriods

                dCellValue = tempAsgnEffort(t)
                dRoundUp = Ceiling(dCellValue, 0.05)
                dRoundDown = Ceiling(dCellValue, -0.05)
                If (nTableType = LEV_RES) Then
                    If (dCellValue - dRoundDown) < (dRoundUp - dCellValue) Then
                        dCellValue = dRoundDown
                    Else
                        dCellValue = dRoundUp
                    End If
                End If
                If Not dCellValue = 0 Then
                    dTotResAlloc = dTotResAlloc + dCellValue


                    If nTableType <> CCL_RES And nTableType <> VAR_RES Then
                        ws1.cells(nCurRow + 2, nCurCol) = dCellValue
                        ws1.cells(nCurRow + 2, nCurCol).NumberFormat = "#,##0.00"

                    End If
                End If

                nCurCol = nCurCol + 1
            Next t

            nCurRow = nCurRow + 1

        End If

    Next r


    ws1.cells(nCurRow + 2, 1) = "   Total"
    ws1.cells(nCurRow + 2, 6).FormulaR1C1 = "=SUM(R" + CStr(nStartRow + 2) + "C:R" + CStr(nStartRow + 2 + m_nNumResources - 1) + "C)"
    ws1.cells(nCurRow + 2, 6).NumberFormat = "#,##0.00"
    ws1.cells(nCurRow + 2, 7).FormulaR1C1 = "=SUM(R" + CStr(nStartRow + 2) + "C:R" + CStr(nStartRow + 2 + m_nNumResources - 1) + "C)"
    ws1.cells(nCurRow + 2, 7).NumberFormat = "#,##0.00"
    ws1.cells(nCurRow + 2, 8).FormulaR1C1 = "=SUM(R" + CStr(nStartRow + 2) + "C:R" + CStr(nStartRow + 2 + m_nNumResources - 1) + "C)"
    ws1.cells(nCurRow + 2, 8).NumberFormat = "$#,##0.00"

    ws1.cells(nCurRow + 2, 9).FormulaR1C1 = "=SUM(R" + CStr(nStartRow + 2) + "C:R" + CStr(nStartRow + 2 + m_nNumResources - 1) + "C)"
    ws1.cells(nCurRow + 2, 9).NumberFormat = "#,##0.00"
    ws1.cells(nCurRow + 2, 10).FormulaR1C1 = "=SUM(R" + CStr(nStartRow + 2) + "C:R" + CStr(nStartRow + 2 + m_nNumResources - 1) + "C)"
    ws1.cells(nCurRow + 2, 10).NumberFormat = "#,##0.00"
    ws1.cells(nCurRow + 2, 11).FormulaR1C1 = "=SUM(R" + CStr(nStartRow + 2) + "C:R" + CStr(nStartRow + 2 + m_nNumResources - 1) + "C)"
    ws1.cells(nCurRow + 2, 11).NumberFormat = "$#,##0.00"

    ' Bottom line (Res Summary)
    ws1.Range(ws1.cells(nCurRow + 2, 1), ws1.cells(nCurRow + 2, FISCAL_MONTH_SCHED1_START_COL - 2)).Interior.Color = RGB(15, 37, 63)
    ws1.Range(ws1.cells(nCurRow + 2, 1), ws1.cells(nCurRow + 2, FISCAL_MONTH_SCHED1_START_COL - 2)).Font.Color = vbWhite
    ws1.Range(ws1.cells(nCurRow + 2, 1), ws1.cells(nCurRow + 2, FISCAL_MONTH_SCHED1_START_COL - 2)).Font.Bold = True


    ' Bottom line (Detailed Schedule)
    For j = DET_SCHED1_START_COL To (DET_SCHED1_START_COL + m_nNumPeriods - 1)
        ws1.cells(nCurRow + 2, j).FormulaR1C1 = "=SUM(R" + CStr(nStartRow + 2) + "C" + CStr(j) + ":R" + CStr(nCurRow + 1) + "C" + CStr(j) + ")"
        ws1.cells(nCurRow + 2, j).NumberFormat = "#,##0.00"
        ws1.cells(nCurRow + 2, j).Interior.Color = RGB(15, 37, 63)
        ws1.cells(nCurRow + 2, j).Font.Color = vbWhite
    Next

    ' Conditional formatting for the Detailed Schedule Grid
    Call AddConditionalFormatting(ws1, nStartRow + 2, DET_SCHED1_START_COL, nStartRow + 2 + m_nNumResources - 1, DET_SCHED1_START_COL + m_nNumPeriods, bReverse)


    If nTableType = VAR_RES Then
        ws1.Range(ws1.cells(Res_Variance_Start_Row + 2, DET_SCHED1_START_COL + m_nFirstDetScheduleWeekInFiscal - 1), _
                  ws1.cells(Res_Variance_Start_Row + 2 + m_nNumResources, DET_SCHED1_START_COL + m_nFirstDetScheduleWeekInFiscal - 1 + m_nPeriodsToCopy - 1)).Value = _
                  "=R[" + CStr(CCL_Spreadsheet_Start_Row - Res_Variance_Start_Row) + "]C-R[" + CStr(LEV_RES_PLAN_START_ROW - Res_Variance_Start_Row) + "]C"

        ws1.Range(ws1.cells(Res_Variance_Start_Row + 2, DET_SCHED1_START_COL + m_nFirstDetScheduleWeekInFiscal - 1), _
                  ws1.cells(Res_Variance_Start_Row + 2 + m_nNumResources, DET_SCHED1_START_COL + m_nFirstDetScheduleWeekInFiscal - 1 + m_nPeriodsToCopy - 1)).NumberFormat = "#,##0.00"
    End If

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Description:
'
'
' Authors:      Cesar Klahr

'
' Comment
' --------------------------------------------------------------
' Initial version
'
Public Sub MI_Resource_Plan2()
    ' Sub will export timephased resource data (work, cost) into Microsoft Excel worksheet

    ' Define timescale unit. Can be one of the following PjTimescaleUnit constants:
    ' pjTimescaleYears, pjTimescaleQuarters, pjTimescaleMonths, pjTimescaleWeeks,pjTimescaleDays, pjTimescaleHours, pjTimescaleMinutes
    Dim tsu As PjTimescaleUnit
    Dim TSVWork As TimeScaleValues

    Dim pj As Project
    Dim sMacroCode As String
    Dim asPrjInfo() As String
    Dim bNotCorrect As Boolean

    Dim sFiscalStartDate As String
    Dim sFiscalEndDate As String
    Dim nLastRowFiscalCalendar As Long





    '//////////////// EARLY BINDING DECLARATIONS////////////////
    'Dim obj As OLEObject
    'Dim ws As Excel.Worksheet
    'Dim rngDataRange As Range
    'Dim rngCell As Range
    '////////////////////////////////////////////////////////////

    '//////////////// LATE BINDING DECLARATIONS////////////////
    Dim obj As Object
    Dim ws As Object
    Dim rngDataRange As Object
    Dim rngCell As Object
    '////////////////////////////////////////////////////////////



    bNotCorrect = True
    While bNotCorrect
        m_sFiscalYear = InputBox("Please enter Fiscal Year to use: ")
        If m_sFiscalYear = "" Then
            End
        End If

        If Not IsNumeric(m_sFiscalYear) Then
            MsgBox "Please enter a valid number for the year"
        Else
            bNotCorrect = False
        End If
    Wend

    Set pj = ActiveProject
    tsu = pjTimescaleWeeks

    asPrjInfo = Split(pj.Name, " ", 3)
    m_sProjectCode = Trim(asPrjInfo(0))


    If pj.Resources.Count > 0 Then
        m_nNumResources = pj.ResourceCount
        m_sProjStartDate = pj.ProjectSummaryTask.Start
        m_sPrjEndDate = pj.ProjectSummaryTask.Finish

        '''''m_sFiscalFileName = "2014 MIS Portfolio Dashboard.xlsm"
        m_sFiscalFileName = "Fiscal Calendar.xlsx"
        m_sFiscalCalendarTabName = "Fiscal Calendar"

        Dim m_sResTrackingFileNamePrefix As String
        Dim m_sResTrackingTabName As String

        m_sResTrackingFileNamePrefix = "Resource Tracking"
        m_sResTrackingTabName = "Resources Allocated"

        m_sFiscalCalFolder = "\\mffilec\datashare\MIS Delivery\Delivery Planning\Portfolio Planning\" + Trim(m_sFiscalYear) + "\"
        m_sCCResAllocFolder = "\\mffilec\datashare\MIS Resourcing\" + Trim(m_sFiscalYear) + "\"

        '''''''m_sCCResAllocFolder = "c:\users\cklahr\Documents\Data\" (copiar en local folder or acceder sin abrir)
        '''''''m_sFiscalCalFolder = "c:\users\cklahr\Documents\Data\" (copiar en local folder or acceder sin abrir)



        m_sFiscalMonthFilePath = m_sFiscalCalFolder
        ''''m_sFiscalMonthFilePath = Trim(MyDocuments) + "\" + Trim(m_sFiscalYear) + "\"     ' eg. "C:\Users\cklahr\Documents\2014\"


        m_sResTrackingFilePath = m_sCCResAllocFolder
        '''''m_sResTrackingFilePath = Trim(MyDocuments) + "\" Trim(m_sFiscalYear) + "\"  ' eg. "C:\Users\cklahr\Documents\2014\"


        '''''''''Set m_xlApp = New Excel.Application
        Set m_xlApp = CreateObject("Excel.Application")


        m_xlApp.Visible = True

        m_xlApp.ReferenceStyle = xlA1LB        ' as opposed to xlR1C1

        'm_xlApp.Calculation = xlCalculationManual

        ' On Error Resume Next

        ''''''''''Dim CalcMode As Long
        '''''''''''CalcMode = Application.Calculation

        ''''''''''''''''''''''''''''''''''
        'Call error handler for cleanup!!!
        ''''''''''''''''''''''''''''''''''!

        m_xlApp.DisplayAlerts = False
        'm_xlApp.Calculation = xlCalculationManual
        m_xlApp.ScreenUpdating = False



        On Error Resume Next
        Set m_wbFC = m_xlApp.Workbooks.Open(m_sFiscalMonthFilePath + m_sFiscalFileName, , True)
        If (m_wbFC Is Nothing) Then
            Set pj = Nothing
            Set m_xlApp = Nothing
            MsgBox ("Can't open " + m_sFiscalFileName + ". Please check and try again later")
            End
        End If

        Set m_wsFC = m_wbFC.Sheets(m_sFiscalCalendarTabName)
        If (m_wsFC Is Nothing) Then
            Set pj = Nothing
            Set m_xlApp = Nothing
            MsgBox ("Can't find Fiscal Calendar tab. Please check and try again later")
            End
        End If

        ' Validate this is the right calendar
        If (m_sFiscalYear <> CStr(Year(m_wsFC.Range("L10")))) Then
            Set pj = Nothing
            Set m_xlApp = Nothing
            MsgBox ("The Fiscal Calendar file contains dates from a different year. Please check and try again later")
            End

        End If

        On Error Resume Next
        Set m_wbResTracking = m_xlApp.Workbooks.Open(m_sResTrackingFilePath + m_sResTrackingFileNamePrefix + " " + Trim(m_sFiscalYear) + ".xlsm", , True)
        If (m_wbResTracking Is Nothing) Then
            Set pj = Nothing
            Set m_xlApp = Nothing
            MsgBox ("Can't open " + m_sResTrackingFileNamePrefix + " " + Trim(m_sFiscalYear) + ".xlsm" + ". Please check and try again later")
            End
        End If

        Set m_wsResAllocated = m_wbResTracking.Sheets(m_sResTrackingTabName)
        If (m_wsResAllocated Is Nothing) Then
            Set pj = Nothing
            Set m_xlApp = Nothing
            MsgBox ("Can't find " + m_sResTrackingTabName + " tab" + ". Please check and try again later")
            End
        End If

        sFiscalStartDate = m_wsFC.cells(4, 12)

        nLastRowFiscalCalendar = m_wsFC.Range("L65536").End(xlUpLB).row
        sFiscalEndDate = m_wsFC.cells(nLastRowFiscalCalendar, 12)

        m_nLastFiscalWeekNum = m_wsFC.cells(nLastRowFiscalCalendar, 12).Offset(, -2)        ' Column L is #12


        If CDate(m_sPrjEndDate) < CDate(sFiscalStartDate) Then

            MsgBox ("The project end date is prior to the start of the " + m_sFiscalFileName)

            m_wbFC.Close savechanges:=False
            m_wbResTracking.Close savechanges:=False

            Set pj = Nothing
            Set m_xlApp = Nothing

            ''''''''''''''''''''''''''''''''''
            'Call error handler for cleanup!!!
            ''''''''''''''''''''''''''''''''''!
            End
        ElseIf CDate(m_sProjStartDate) > CDate(sFiscalEndDate) Then
            MsgBox ("The project start date is later than the end of the " + m_sFiscalFileName)

            m_wbFC.Close savechanges:=False
            m_wbResTracking.Close savechanges:=False

            Set pj = Nothing
            Set m_xlApp = Nothing

            ''''''''''''''''''''''''''''''''''
            'Call error handler for cleanup!!!
            ''''''''''''''''''''''''''''''''''!
        End If

        Set m_wbResPlan = m_xlApp.Workbooks.Add
        m_wbResPlan.Application.WindowState = xlMaximizedLB
        m_wbResPlan.Title = pj.Title

        Set m_wsResPlan = m_wbResPlan.Worksheets.Add
        m_wsResPlan.Name = "ResPlan"
        m_wsResPlan.cells.Font.Name = "Calibri"
        m_wsResPlan.cells.Font.Size = 8
        m_xlApp.ActiveWindow.DisplayGridlines = False        ' Remove Gridlines

        Call TurnCalculationOnAndOff(m_xlApp)

        Application.DisplayAlerts = False
        For Each ws In m_wbResPlan.Sheets
            If ws.Name <> "ResPlan" Then ws.Delete
        Next
        Application.DisplayAlerts = True

        m_wsResPlan.cells(1, 1) = pj.Name & " - Resource Plan"
        m_wsResPlan.cells(1, 1).Font.Bold = True
        m_wsResPlan.cells(1, 1).ColumnWidth = 40
        m_wsResPlan.cells(1, 1).Font.Size = 12

        m_wsResPlan.cells(2, 1) = "Last Updated on: " + CStr(Format(Date, "Long Date"))

        ' Need to know how many periods are there (I know there is at least one resource in the roster)
        Set TSVWork = pj.Resources(1).TimeScaleData(m_sProjStartDate, m_sPrjEndDate, Type:=pjResourceTimescaledWork, TimescaleUnit:=tsu)
        m_nNumPeriods = TSVWork.Count

        'Create button to refresh the CCL table in the future
        Set obj = m_wsResPlan.OLEObjects.Add(ClassType:="Forms.CommandButton.1", Link:=False, DisplayAsIcon:=False, Left:=300, Top:=35, Width:=220, Height:=30)

        ' Set the button's name and caption
        obj.Name = "RefreshFromCCLTab"
        obj.Object.Caption = "RFCCLT"
        obj.Object.Caption = "REFRESH FROM CCL SPREADSHEET"
        obj.Object.BackColor = &H808000
        obj.Object.ForeColor = &HC0FFFF
        obj.Object.Font.Bold = True
        obj.Object.Font.Size = 14

        'Macro text
        sMacroCode = "Sub RefreshFromCCLTab_click" & vbCrLf
        sMacroCode = sMacroCode & "    MsgBox "" You won a Cruise for the Bahamas!""" & vbCrLf
        sMacroCode = sMacroCode & "End Sub"

        'Add macro at the end of the sheet module
        With m_wbResPlan.VBProject.VBComponents(m_wsResPlan.CodeName).CodeModule
            .insertlines .CountOfLines + 1, sMacroCode
        End With


        If m_nNumPeriods > 0 Then
            ''If (pj.StartWeekOn = pjSaturday) Then
            ''    m_nDistance2NextFriday = 6
            ''Else
            ''    m_nDistance2NextFriday = 6 - pj.StartWeekOn
            ''End If

            '' m_nDistance2NextFriday = 8 - Weekday(pj.StartWeekOn, vbFriday)
            If (Weekday(TSVWork(1).StartDate, vbFriday) > 1) Then
                m_nDistance2NextFriday = 8 - Weekday(TSVWork(1).StartDate, vbFriday)
            Else
                m_nDistance2NextFriday = 0        'Its already Friday
            End If




            Set m_list = New VbaList
            Set m_list = m_list.CreateInstance

            ' Create linked list with task ids that have been selected (m_list)
            Call LoadSelectedTasks(pj)





            ''''Adjust number of resources to those who have assignments and should be listed!!! ONLY LIST RESOURCES WITH ASSIGNMENTES!!!!!
            Dim nRes As Long

            nRes = 0
            Dim Asgn As Assignment
            Dim r As Long

            For r = 1 To m_nNumResources
                For Each Asgn In pj.Resources(r).Assignments

                    ' If the assignment is for a task in the selection, THEN ADD THE RESOURCE AND TIMEPHASED DATA!!!
                    If m_list.Exists(Asgn.Task.ID) Then

                        nRes = nRes + 1
                        Exit For        ' Dont care if there is more than 1 assignment

                    End If
                Next Asgn
            Next r

            m_nNumResources = nRes


            MSP_Res_Plan_Start_Row = LEV_RES_PLAN_START_ROW + m_nNumResources + SEPARATOR
            CCL_Spreadsheet_Start_Row = MSP_Res_Plan_Start_Row + m_nNumResources + SEPARATOR
            Res_Variance_Start_Row = CCL_Spreadsheet_Start_Row + m_nNumResources + SEPARATOR




            Call Draw_tables(pj, m_wsResPlan, LEV_RES_PLAN_START_ROW, TSVWork, LEV_RES, tsu)
            Call Draw_tables(pj, m_wsResPlan, MSP_Res_Plan_Start_Row, TSVWork, MSP_RES, tsu)
            '@@@@@@@@@@@@@@@@ Call Draw_tables(pj, m_wsResPlan, CCL_Spreadsheet_Start_Row, TSVWork, CCL_RES, tsu)
            '@@@@@@@@@@@@@@@@ Call Draw_tables(pj, m_wsResPlan, Res_Variance_Start_Row, TSVWork, VAR_RES, tsu)




            Set m_list = Nothing


        Else
            MsgBox "No periods found in the project!"
        End If

    Else
        MsgBox "No resources has been assigned to the project yet!"
        End
    End If

    Call CopyCCLAllocations(m_wsResPlan)
    Call CleanUPCCLAllocations(m_wsResPlan)


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'ADD CODE TO DELETE ANYTHING THAT LOOK GARBAGE IN THE RES_RACKING_TABLE!!!!!
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' Nee to reapply format as Pastevalues is not working!
    m_wsResPlan.Range(m_wsResPlan.cells(CCL_Spreadsheet_Start_Row + 2, DET_SCHED1_START_COL), _
                      m_wsResPlan.cells(CCL_Spreadsheet_Start_Row + 2 + m_nNumResources - 1, DET_SCHED1_START_COL + m_nNumPeriods)).NumberFormat = "#,##0.00"

    'On Error Resume Next
    'Application.Calculation = xlCalculationAutomatic
    ''''''''''''''''''''''''''''''''''
    'Call error handler for cleanup!!!
    ''''''''''''''''''''''''''''''''''!

    m_wsResPlan.Range("C7").Select
    m_xlApp.ActiveWindow.FreezePanes = True

    Call TurnCalculationOnAndOff(m_xlApp)

    m_wbFC.Close savechanges:=False
    m_wbResTracking.Close savechanges:=False



    m_xlApp.DisplayAlerts = True
    m_xlApp.Calculation = xlCalculationAutomatic
    m_xlApp.ScreenUpdating = True


    Set pj = Nothing
    Set m_xlApp = Nothing
    Set m_wbResPlan = Nothing
    Set m_wsResPlan = Nothing
    Set TSVWork = Nothing

    'and finally display a message that we are finished
    AppActivate "Microsoft Project"

    'm_xlApp.Visible = True

    MsgBox "Resource Plan macro complete!"

    ' AppActivate "Microsoft Excel"
End Sub














