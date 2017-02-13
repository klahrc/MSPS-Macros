Attribute VB_Name = "MI_Task_Hierarchy3"

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
' Description: This macro will export tasks to excel while keeping the task hierarchy
' Start Date: 23-Jun-2013
'
' Pending
' *******
' Improve error handling
' Release Objects
' Add precedences
' Allow to choose which baseline to use
' Add Summarized Cost info in the report
'
' Completed
' *********
' Show if activity didnt start in AB (red dot). OK
' Higlight in red font if plan is not matching baseline. OK
' Show activities starting in next 2 weeks in yellow. OK
' Show green for activities in progress. OK
' Test if there is no ProjectSummaryTask!

Option Explicit

'Public gbIncludeResInfo As Boolean
'Public gbProcessAllTasks As Boolean


'//////////////// EARLY BINDING DECLARATIONS////////////////
'Dim m_rngRow As Excel.Range                                                                ' Row Index
'Dim m_rngCol As Excel.Range                                                                ' Column Index
'////////////////////////////////////////////////////////////

'//////////////// LATE BINDING DECLARATIONS////////////////
Dim m_rngRow As Object        ' Row Index
Dim m_rngCol As Object        ' Column Index
'////////////////////////////////////////////////////////////

Dim m_WoD As String        ' Work or Days in the Gantt Chart?

Sub tideUp()
    Dim tsk As Task

    For Each tsk In ActiveProject.Tasks
        If Not tsk Is Nothing Then
            If tsk.Summary = True Then
                If tsk.PercentComplete >= 100 Then
                    tsk.OutlineHideSubTasks
                Else
                    tsk.OutlineShowSubTasks


                End If
            End If
        End If
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
Private Sub dwn(i As Integer)
    Set m_rngRow = m_rngRow.Offset(i, 0)
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
Private Sub rgt(i As Integer)
    Set m_rngCol = m_rngCol.Offset(0, i)
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
Public Sub MI_TH()
    '//////////////// EARLY BINDING DECLARATIONS////////////////
    '    Dim xlApp As Excel.Application
    '    Dim wbWkBk1 As Excel.Workbook
    '    Dim wsSheet1 As Excel.Worksheet
    '    Dim rngLastCol As Excel.Range                                                        ' Last column before Gantt Chart
    '    Dim rngTasKRow As Excel.Range                                                        ' Task Row (need to go back to paint status)
    '////////////////////////////////////////////////////////////

    '//////////////// LATE BINDING DECLARATIONS////////////////
    Dim xlApp As Object
    Dim wbWkBk1 As Object
    Dim wsSheet1 As Object
    Dim rngLastCol As Object        ' Last column before Gantt Chart
    Dim rngTasKRow As Object        ' Task Row (need to go back to paint status)
    '////////////////////////////////////////////////////////////

    Dim pj As Project        ' Project variable
    Dim t As Task        ' Task variable
    Dim Asgn As Assignment        ' Asgn variable

    Dim iOutlineCol As Integer        ' Store number of columns needed for outline levels
    Dim iColumns As Integer        ' Index for current column
    Dim iTotTasks As Integer        ' Total number of tasks in the project

    Dim d As Date        ' Temporary variable to hold dates
    Dim dGntStart, dGntFinish As Date        ' Start and Finish dates for Gantt Chart

    Dim iWeekDayGanttStart As Integer        ' WeekDay of the Start Date in the Gantt Chart
    Dim iNumOfWeeksInGanttChart As Integer        ' Number of weeks in Gantt Chart
    Dim iX1GanttCoord, iY1GanttCoord As Integer
    Dim iX2GanttCoord, iY2GanttCoord As Integer
    Dim i, p        ' Index variable (multipurpose)
    Dim iReply As Integer        ' Store user's reply

    'Dim bAddResInfo As Boolean

    Dim oProgress As IPRogressBar        'Progress Bar control


    Dim TSVActualWork As TimeScaleValues        ' Will hold Timescale data for Actual Work
    Dim TSVPlannedWork As TimeScaleValues        ' Will hold Timescale data for Planned Work

    ' Define timescale unit. Can be one of the following PjTimescaleUnit constants:
    ' pjTimescaleYears, pjTimescaleQuarters, pjTimescaleMonths, pjTimescaleWeeks,pjTimescaleDays, pjTimescaleHours, pjTimescaleMinutes
    Dim TimescaleUnit As PjTimescaleUnit

    TimescaleUnit = pjTimescaleWeeks

    '    iReply = MsgBox(Prompt:="Do you want to include resource information in the report?", _
         '                    Buttons:=vbYesNoCancel, title:="Task Hierarchy")
    '
    '
    '    If iReply = vbYes Then
    '        bAddResInfo = True
    '    ElseIf iReply = vbNo Then
    '        bAddResInfo = False
    '    Else        'Cancelled (VbCancel)
    '        End
    '    End If

    '********************************************************
    m_WoD = "W"        ' Simulate until we have a form!
    '********************************************************

    Set pj = ActiveProject

    If pj.Tasks.Count = 0 Then

        MsgBox ("No Tasks has been added to the project yet!")
        End

    End If


    ' Show hourglass
    Load frmPleaseWait
    frmPleaseWait.Caption = "Please wait..."
    frmPleaseWait.Show False
    DoEvents

    ''' dGntStart = ActiveProject.ProjectStart
    ''' dGntFinish = ActiveProject.ProjectFinish
    dGntStart = ActiveProject.ProjectSummaryTask.Start
    dGntFinish = ActiveProject.ProjectSummaryTask.Finish

    ''' iNumOfWeeksInGanttChart = Int((dGntFinish - dGntStart) / 7) + 2
    iNumOfWeeksInGanttChart = DateDiff("ww", dGntStart, dGntFinish) + 2        ' Adding 1 extra week to show there is nothing planned for the last week


    Set xlApp = CreateObject("Excel.Application")        ' Create a new instance of Excel
    xlApp.Visible = False
    xlApp.ReferenceStyle = xlA1LB        ' as opposed to xlR1C1

    Set wbWkBk1 = xlApp.Workbooks.Add        ' Adding a new workbook

    wbWkBk1.Application.WindowState = xlMinimizedLB


    xlApp.DisplayAlerts = False
    xlApp.Calculation = xlCalculationManual
    xlApp.ScreenUpdating = False


    Set wsSheet1 = wbWkBk1.Worksheets.Add        ' Adding a new spreadsheet
    wsSheet1.Name = Left("TH - " + ActiveProject.Name, 31)

    xlApp.ActiveWindow.DisplayGridlines = False        ' Remove Gridlines

    With wsSheet1.cells.Font        ' Set font properties
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 10
    End With

    Dim lNumTasks As Long

    lNumTasks = 0



    ''''''''''''''''''''''''''''''''''''''''''''' Add Columns and configure the Detailed Report Section '''''''''''''''''''''''''''''''''''''''''''''''
    'Calculating columns needed for outline levels in the spreadsheet
    iOutlineCol = 0        ' I'll use first column for 'row type'


    '''''''''frmTasksResourcesChoices.Hide
    ''''frmToolbox.Hide


    If gbProcessAllTasks Then
        'Expand all rows!
        pj.ProjectSummaryTask.OutlineShowAllTasks

        SelectAll
    End If



    ''''For Each t In ActiveProject.Tasks
    For Each t In ActiveSelection.Tasks
        If Not t Is Nothing Then
            If t.OutlineLevel > iOutlineCol Then
                iOutlineCol = t.OutlineLevel
            End If


            lNumTasks = lNumTasks + 1
        End If
    Next t

    'Set Range to write to first cell
    Set m_rngRow = xlApp.ActiveCell        ' At the beginning ActiveCell is A1
    m_rngRow.ColumnWidth = 1
    m_rngRow = "Filename: " & ActiveProject.Name + " - Date: " + Format(Now(), "ddd dd-mmm-yyyy") + " - Task Hierarchy"
    m_rngRow.Font.Bold = True

    dwn 1
    m_rngRow = "OutlineLevel>>>"

    ' Column labels
    dwn 1
    Set m_rngCol = m_rngRow.Offset(0, 1)
    m_rngCol = "%C"
    m_rngCol.Font.Bold = True
    m_rngCol.Font.Underline = True
    m_rngCol.ColumnWidth = 3

    rgt 1
    m_rngCol = "%P"
    m_rngCol.Font.Bold = True
    m_rngCol.Font.Underline = True
    m_rngCol.ColumnWidth = 3

    rgt 2
    'Label rest of columns
    For iColumns = 2 To (iOutlineCol + 2)
        Set m_rngCol = m_rngRow.Offset(0, iColumns + 1)
        m_rngCol = iColumns - 2
        m_rngCol.ColumnWidth = 3
        m_rngCol.HorizontalAlignment = xlLeft
        m_rngCol.Font.Bold = True
        m_rngCol.Font.Underline = True
    Next

    rgt 1
    m_rngCol = "Resource Name"
    m_rngCol.ColumnWidth = 24
    m_rngCol.Font.Bold = True
    m_rngCol.Font.Underline = True

    rgt 1
    m_rngCol = "P.W.(h)"
    m_rngCol.ColumnWidth = 7
    m_rngCol.Font.Bold = True
    m_rngCol.Font.Underline = True
    m_rngCol.HorizontalAlignment = xlRight

    rgt 1
    m_rngCol = "B.W.(h)"
    m_rngCol.ColumnWidth = 7
    m_rngCol.Font.Bold = True
    m_rngCol.Font.Underline = True
    m_rngCol.HorizontalAlignment = xlRight

    rgt 1
    m_rngCol = "A.W.(h)"
    m_rngCol.ColumnWidth = 7
    m_rngCol.Font.Bold = True
    m_rngCol.Font.Underline = True
    m_rngCol.HorizontalAlignment = xlRight

    rgt 1
    m_rngCol = "R.W.(h)"
    m_rngCol.ColumnWidth = 7
    m_rngCol.Font.Bold = True
    m_rngCol.Font.Underline = True
    m_rngCol.HorizontalAlignment = xlRight

    rgt 1
    m_rngCol = "P. Start"
    m_rngCol.ColumnWidth = 8
    m_rngCol.Font.Bold = True
    m_rngCol.Font.Underline = True
    m_rngCol.HorizontalAlignment = xlRight

    rgt 1
    m_rngCol = "P. Finish"
    m_rngCol.ColumnWidth = 8
    m_rngCol.Font.Bold = True
    m_rngCol.Font.Underline = True
    m_rngCol.HorizontalAlignment = xlRight

    rgt 1
    m_rngCol = "B. Start"
    m_rngCol.ColumnWidth = 8
    m_rngCol.Font.Bold = True
    m_rngCol.Font.Underline = True
    m_rngCol.HorizontalAlignment = xlRight

    rgt 1
    m_rngCol = "B. Finish"
    m_rngCol.ColumnWidth = 8
    m_rngCol.Font.Bold = True
    m_rngCol.Font.Underline = True
    m_rngCol.HorizontalAlignment = xlRight

    rgt 1
    m_rngCol = "A. Start"
    m_rngCol.ColumnWidth = 8
    m_rngCol.Font.Bold = True
    m_rngCol.Font.Underline = True
    m_rngCol.HorizontalAlignment = xlRight

    rgt 1
    m_rngCol = "A. Finish"
    m_rngCol.ColumnWidth = 8
    m_rngCol.Font.Bold = True
    m_rngCol.Font.Underline = True
    m_rngCol.HorizontalAlignment = xlRight

    rgt 1
    m_rngCol = "AB"
    m_rngCol.ColumnWidth = 2
    m_rngCol.Font.Bold = True
    m_rngCol.Font.Underline = True

    rgt 1
    m_rngCol = "OU"
    m_rngCol.ColumnWidth = 2
    m_rngCol.Font.Bold = True
    m_rngCol.Font.Underline = True

    rgt 1
    m_rngCol = "PU"
    m_rngCol.ColumnWidth = 2
    m_rngCol.Font.Bold = True
    m_rngCol.Font.Underline = True

    rgt 1
    m_rngCol.ColumnWidth = 1

    Set rngLastCol = m_rngCol        ' Save last column before Gantt

    '''''''''''''''''''''''''''''''''''''''''''''''''' Add Columns and configure the Gannt Chart '''''''''''''''''''''''''''''''''''''''''''''''''''''

    '''' iWeekDayGanttStart = Weekday(dGntStart)                                              ' Weekday returns 1,7 (Sunday to Saturday)
    '''' d = dGntStart + 6 - iWeekDayGanttStart                                               ' Align to Fridays 6=Friday! //////////////////

    d = dGntStart
    If (Weekday(dGntStart, vbFriday) > 1) Then
        d = dGntStart + 8 - Weekday(dGntStart, vbFriday)        ' Align to Fridays! //////////////////
    End If



    For i = 1 To iNumOfWeeksInGanttChart
        ' Add columns for Gantt chart
        m_rngCol.EntireColumn.Offset(0, 1).Insert

        If (m_WoD) = "D" Then        ' When showing Days in the Gantt blocks
            m_rngCol.EntireColumn.Offset(0, 1).ColumnWidth = 2
        Else        ' Otherwise we need more room to show Work
            m_rngCol.EntireColumn.Offset(0, 1).ColumnWidth = 4
        End If

        ' Add dates for Gantt chart
        m_rngCol.Offset(0, 1) = d
        m_rngCol.Offset(0, 1).NumberFormat = "dd"
        m_rngCol.Offset(0, 1).Interior.Color = RGB(204, 255, 255)
        m_rngCol.Offset(0, 1).Font.Name = "Candara"
        m_rngCol.Offset(0, 1).Font.Size = 9

        m_rngCol.Offset(-1, 1) = d
        m_rngCol.Offset(-1, 1).NumberFormat = "mm"
        m_rngCol.Offset(-1, 1).Interior.Color = RGB(204, 255, 255)
        m_rngCol.Offset(-1, 1).Font.Name = "Candara"
        m_rngCol.Offset(-1, 1).Font.Size = 10

        m_rngCol.Offset(-2, 1) = d
        m_rngCol.Offset(-2, 1).NumberFormat = "yy"
        m_rngCol.Offset(-2, 1).Interior.Color = RGB(204, 255, 255)
        m_rngCol.Offset(-2, 1).Font.Name = "Candara"
        m_rngCol.Offset(-2, 1).Font.Size = 11

        d = d + 7
        rgt 1
    Next

    '''''''''''''''''''''''''''''''''''''''''''''''''' Logic for the Project Summary Task only'''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set t = pj.ProjectSummaryTask

    If Not t Is Nothing Then
        If t.Active Then
            dwn 1
            Set m_rngCol = m_rngRow.Offset(0, 1)
            m_rngCol = t.PercentComplete
            m_rngCol.Font.Size = 8

            rgt 1
            m_rngCol = t.PhysicalPercentComplete
            m_rngCol.Font.Size = 8

            If (t.PercentComplete <> t.PhysicalPercentComplete) Then        ' Highlight in red font if %C differs from PhysicalPercentComplete
                m_rngCol.Font.Color = RGB(255, 0, 0)
                m_rngCol.Font.Bold = True
                m_rngCol.Offset(0, -1).Font.Color = RGB(255, 0, 0)
                m_rngCol.Offset(0, -1).Font.Bold = True
            End If


            Set m_rngCol = m_rngRow.Offset(0, t.OutlineLevel + 3)
            m_rngCol = t.Name
            If (t.Manual) Then        ' If task mode is manual highlight in red font
                m_rngCol = "(M) " + t.Name
                m_rngCol.Font.Color = RGB(255, 0, 0)
            End If
            m_rngCol.Font.Bold = True
            wsSheet1.cells(m_rngCol.row, 1) = "S" + CStr(t.OutlineLevel)        ' Indicate it's a Summary Task and outlevel
            wsSheet1.cells(m_rngCol.row, 1).Font.Name = "Kartika"
            wsSheet1.cells(m_rngCol.row, 1).Font.Size = 8

            rgt (iOutlineCol - m_rngCol.Column) + 6
            If (t.Duration <> 0) Then
                m_rngCol = (t.Work / 60)        ' Task work is stored is in minutes
                m_rngCol.NumberFormat = "#,##0.0"
                m_rngCol.Font.Bold = True

                rgt 1
                m_rngCol = (t.BaselineWork / 60)        ' Task work is stored is in minutes
                m_rngCol.NumberFormat = "#,##0.0"
                m_rngCol.Font.Bold = True

                rgt 1
                m_rngCol = (t.ActualWork / 60)        ' Task work is stored is in minutes
                m_rngCol.NumberFormat = "#,##0.0"
                m_rngCol.Font.Bold = True

                rgt 1
                m_rngCol = (t.RemainingWork / 60)        ' Task work is stored is in minutes
                m_rngCol.NumberFormat = "#,##0.0"
                m_rngCol.Font.Bold = True
            Else
                rgt 3
            End If

            rgt 1
            m_rngCol = DateValue(t.Start)        ' Planned Start
            m_rngCol.Font.Bold = True
            m_rngCol.NumberFormat = "mm/dd/yy"

            rgt 1
            m_rngCol = DateValue(t.Finish)        ' Planned Finsish
            m_rngCol.Font.Bold = True
            m_rngCol.NumberFormat = "mm/dd/yy"

            rgt 1
            If IsDate(t.BaselineStart) Then        ' If the project has been baselined
                m_rngCol = DateValue(t.BaselineStart)        ' Show Baseline Start
                m_rngCol.NumberFormat = "mm/dd/yy"
                If (DateValue(t.Start) > DateValue(t.BaselineStart)) Then
                    m_rngCol.Font.Color = RGB(255, 0, 0)
                    m_rngCol.Offset(0, -2).Font.Color = RGB(255, 0, 0)
                End If
            Else
                m_rngCol = t.BaselineStart        ' N/A (No Baseline Start)
                m_rngCol.Font.Color = RGB(255, 0, 0)        ' Highlight in red font
                m_rngCol.HorizontalAlignment = xlRight
            End If
            m_rngCol.Font.Bold = True

            rgt 1
            If IsDate(t.BaselineFinish) Then        ' If the project has been baselined
                m_rngCol = DateValue(t.BaselineFinish)        ' Show Baseline Finish
                m_rngCol.NumberFormat = "mm/dd/yy"
                If (DateValue(t.Finish) > DateValue(t.BaselineFinish)) Then
                    m_rngCol.Font.Color = RGB(255, 0, 0)
                    m_rngCol.Offset(0, -2).Font.Color = RGB(255, 0, 0)
                End If
            Else
                m_rngCol = t.BaselineFinish        ' N/A (No Baseline Finish)
                m_rngCol.Font.Color = RGB(255, 0, 0)        ' Highlight in red font
                m_rngCol.HorizontalAlignment = xlRight
            End If
            m_rngCol.Font.Bold = True

            rgt 1
            If IsDate(t.ActualStart) Then        ' If the task has started
                m_rngCol = DateValue(t.ActualStart)        ' Show actual Start
                m_rngCol.NumberFormat = "mm/dd/yy"
            Else
                m_rngCol = t.ActualStart        ' N/A (The task hasn't started yet)
                m_rngCol.HorizontalAlignment = xlRight
            End If
            m_rngCol.Font.Bold = True

            rgt 1
            If IsDate(t.ActualFinish) Then        ' If the task has finished
                m_rngCol = DateValue(t.ActualFinish)        ' Show actual Finish
                m_rngCol.NumberFormat = "mm/dd/yy"
            Else
                m_rngCol = t.ActualFinish        ' N/A
                m_rngCol.HorizontalAlignment = xlRight
            End If
            m_rngCol.Font.Bold = True

            rgt 1
            m_rngCol = Chr$(149)        ' Ahead or Behind schedule?  - 149 is ASCII for the dot
            m_rngCol.Font.Size = 11
            m_rngCol.Font.Color = RGB(0, 128, 0)        ' Show Green dot by default
            If (t.PercentComplete = 0 And DateValue(t.Start) < DateValue(Now())) Then        ' Task should have started but it didn't
                m_rngCol.Font.Color = RGB(255, 0, 0)        ' Red dot
                m_rngCol.Offset(0, -6).Font.Color = RGB(255, 0, 0)
            ElseIf (t.PercentComplete < 100) Then
                If (IsDate(t.BaselineStart) And IsDate(t.BaselineFinish)) Then        ' If the project is baselined
                    If ((DateValue(t.Finish) > DateValue(t.BaselineFinish)) Or _
                        (DateValue(t.Start) > DateValue(t.BaselineStart))) Then        ' If there is a slippage
                        m_rngCol.Font.Color = RGB(255, 0, 0)        ' Show Red dot
                    End If
                Else
                    m_rngCol.Font.Color = RGB(255, 0, 0)        ' Show Red dot if there is no baseline info!
                End If
            End If
            m_rngCol.Font.Bold = True
            m_rngCol.HorizontalAlignment = xlCenter

            rgt 1
            m_rngCol = Chr$(149)        ' Over or Under estimates?  - 149 is ASCII for the dot
            m_rngCol.Font.Size = 11
            m_rngCol.Font.Color = RGB(0, 128, 0)        ' Show Green dot by default
            If (IsDate(t.BaselineStart)) Then        ' Perhaps I should just ask if ISNUM (t.BaselineWork)????
                If (t.PercentComplete < 100 And t.Work > t.BaselineWork) Then        ' If AC is greater than BL and the task is not completed
                    m_rngCol.Font.Color = RGB(255, 0, 0)        ' Show Red dot
                End If
            End If
            m_rngCol.Font.Bold = True
            m_rngCol.HorizontalAlignment = xlCenter
        End If
    End If

    '''''''''''''''''''''''''''''''''''''''''''''''''''''' Logic for the other Project Tasks ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    iTotTasks = 0




    ' Remove hourglass
    frmPleaseWait.Hide



    Set oProgress = New ProgressDialogue


    ' Showing progress
    oProgress.Configure "Loading Tasks in TH Report. Please wait...", "", 0, lNumTasks

    oProgress.Show



    Dim lProgress As Long
    lProgress = 0



    ''''    For Each t In ActiveProject.Tasks
    For Each t In ActiveSelection.Tasks
        If Not t Is Nothing Then
            If t.Active Then

                lProgress = lProgress + 1
                oProgress.Value = lProgress
                oProgress.Status = "Added " + Trim(Str(lProgress)) + " Tasks"



                dwn 1
                Set rngTasKRow = m_rngRow        ' Save start row for this task
                Set m_rngCol = m_rngRow.Offset(0, 1)
                m_rngCol = t.PercentComplete
                m_rngCol.Font.Size = 8

                rgt 1
                m_rngCol = t.PhysicalPercentComplete
                m_rngCol.Font.Size = 8
                If (t.PercentComplete <> t.PhysicalPercentComplete) Then
                    m_rngCol.Font.Color = RGB(255, 0, 0)
                    m_rngCol.Offset(0, -1).Font.Color = RGB(255, 0, 0)
                    If (t.Summary) Then
                        m_rngCol.Font.Bold = True
                        m_rngCol.Offset(0, -1).Font.Bold = True
                    End If
                End If
                Set m_rngCol = m_rngRow.Offset(0, t.OutlineLevel + 3)
                m_rngCol = t.Name
                If (t.Manual) Then
                    m_rngCol = "(M) " + t.Name
                    m_rngCol.Font.Color = RGB(255, 0, 0)
                End If
                If t.Summary Then
                    m_rngCol.Font.Bold = True
                    wsSheet1.cells(m_rngCol.row, 1) = "S" + CStr(t.OutlineLevel)        ' Indicate it's a Summary Task and outlevel
                    wsSheet1.cells(m_rngCol.row, 1).Font.Name = "Kartika"
                    wsSheet1.cells(m_rngCol.row, 1).Font.Size = 8
                ElseIf (t.Duration = 0) Then
                    wsSheet1.cells(m_rngCol.row, 1) = "M"
                    wsSheet1.cells(m_rngCol.row, 1).Font.Name = "Kartika"
                    wsSheet1.cells(m_rngCol.row, 1).Font.Size = 8
                Else
                    wsSheet1.cells(m_rngCol.row, 1) = "T" + CStr(t.OutlineLevel)        ' Indicate it's a Task and outlevel
                    wsSheet1.cells(m_rngCol.row, 1).Font.Name = "Kartika"
                    wsSheet1.cells(m_rngCol.row, 1).Font.Size = 8
                End If

                rgt (iOutlineCol - m_rngCol.Column) + 6
                If (t.Duration <> 0) Then
                    m_rngCol = (t.Work / 60)        ' Task work is stored is in minutes
                    m_rngCol.NumberFormat = "#,##0.0"
                    If (t.Summary) Then
                        m_rngCol.Font.Bold = True
                    End If
                    rgt 1
                    m_rngCol = (t.BaselineWork / 60)        ' Task work is stored is in minutes
                    m_rngCol.NumberFormat = "#,##0.0"
                    If (t.Summary) Then
                        m_rngCol.Font.Bold = True
                    End If
                    rgt 1
                    m_rngCol = (t.ActualWork / 60)        ' Task work is stored is in minutes
                    m_rngCol.NumberFormat = "#,##0.0"
                    If (t.Summary) Then
                        m_rngCol.Font.Bold = True
                    End If
                    rgt 1
                    m_rngCol = (t.RemainingWork / 60)        ' Task work is stored is in minutes
                    m_rngCol.NumberFormat = "#,##0.0"
                    If (t.Summary) Then
                        m_rngCol.Font.Bold = True
                    End If
                Else
                    rgt 3
                End If

                rgt 1
                m_rngCol = DateValue(t.Start)        ' Planned Start
                If (t.Summary) Then
                    m_rngCol.Font.Bold = True
                End If
                m_rngCol.NumberFormat = "mm/dd/yy"

                rgt 1
                m_rngCol = DateValue(t.Finish)        ' Planned Finsh
                If (t.Summary) Then
                    m_rngCol.Font.Bold = True
                End If
                m_rngCol.NumberFormat = "mm/dd/yy"

                rgt 1
                If IsDate(t.BaselineStart) Then        ' If the project has been baselined
                    m_rngCol = DateValue(t.BaselineStart)        ' Show Baseline Start
                    m_rngCol.NumberFormat = "mm/dd/yy"
                    If (DateValue(t.Start) > DateValue(t.BaselineStart)) Then
                        m_rngCol.Font.Color = RGB(255, 0, 0)
                        m_rngCol.Offset(0, -2).Font.Color = RGB(255, 0, 0)
                    End If
                Else
                    m_rngCol = t.BaselineStart        ' N/A (No Baseline Start)
                    m_rngCol.Font.Color = RGB(255, 0, 0)        'Highlight in red font
                    m_rngCol.HorizontalAlignment = xlRight
                End If
                If (t.Summary) Then
                    m_rngCol.Font.Bold = True
                End If

                rgt 1
                If IsDate(t.BaselineFinish) Then        ' If the project has been baselined
                    m_rngCol = DateValue(t.BaselineFinish)        ' Show Baseline Finish
                    m_rngCol.NumberFormat = "mm/dd/yy"
                    If (DateValue(t.Finish) > DateValue(t.BaselineFinish)) Then
                        m_rngCol.Font.Color = RGB(255, 0, 0)
                        m_rngCol.Offset(0, -2).Font.Color = RGB(255, 0, 0)
                    End If
                Else
                    m_rngCol = t.BaselineFinish        ' N/A
                    m_rngCol.Font.Color = RGB(255, 0, 0)        ' Highlight in red font
                    m_rngCol.HorizontalAlignment = xlRight
                End If
                If (t.Summary) Then
                    m_rngCol.Font.Bold = True
                End If

                rgt 1
                If IsDate(t.ActualStart) Then        ' If the task has started
                    m_rngCol = DateValue(t.ActualStart)        ' Show actual Start
                    m_rngCol.NumberFormat = "mm/dd/yy"
                Else
                    m_rngCol = t.ActualStart        ' N/A (No Baseline Finish)
                    m_rngCol.HorizontalAlignment = xlRight
                End If
                If (t.Summary) Then
                    m_rngCol.Font.Bold = True
                End If

                rgt 1
                If IsDate(t.ActualFinish) Then        ' If the task has finished
                    m_rngCol = DateValue(t.ActualFinish)        ' Show actual Finish
                    m_rngCol.NumberFormat = "mm/dd/yy"
                Else
                    m_rngCol = t.ActualFinish        ' N/A
                    m_rngCol.HorizontalAlignment = xlRight
                End If
                If (t.Summary) Then
                    m_rngCol.Font.Bold = True
                End If

                rgt 1
                m_rngCol = Chr$(149)        ' Ahead or Behind schedule?  - 149 is ASCII for the dot
                m_rngCol.Font.Size = 11
                m_rngCol.Font.Color = RGB(0, 128, 0)        ' Show Green dot by default

                If (t.PercentComplete = 0 And DateValue(t.Start) < DateValue(Now())) Then        ' Task should have started but it didn't
                    m_rngCol.Font.Color = RGB(255, 0, 0)        ' Red dot
                    m_rngCol.Offset(0, -6).Font.Color = RGB(255, 0, 0)
                ElseIf (t.PercentComplete < 100) Then
                    If (IsDate(t.BaselineStart) And IsDate(t.BaselineFinish)) Then        ' If the project is baselined
                        If ((DateValue(t.Finish) > DateValue(t.BaselineFinish)) Or _
                            (DateValue(t.Start) > DateValue(t.BaselineStart))) Then        ' If there is a slippage
                            m_rngCol.Font.Color = RGB(255, 0, 0)        ' Show Red dot
                        End If
                    Else
                        m_rngCol.Font.Color = RGB(255, 0, 0)        ' Show Red dot if there is no baseline info!
                    End If
                End If


                m_rngCol.Font.Bold = True
                m_rngCol.HorizontalAlignment = xlCenter
                rgt 1
                m_rngCol = Chr$(149)        ' Over or Under estimates?  - 149 is ASCII for the dot
                m_rngCol.Font.Size = 11
                m_rngCol.Font.Color = RGB(0, 128, 0)        ' Show Green dot by default
                If (IsDate(t.BaselineStart)) Then        ' Perhaps I should just ask if ISNUM (t.BaselineWork)????
                    If (t.PercentComplete < 100 And t.Work > t.BaselineWork) Then        ' If AC is greater than BL and the task is not completed
                        m_rngCol.Font.Color = RGB(255, 0, 0)        ' Show Red dot
                    End If
                End If
                m_rngCol.Font.Bold = True
                m_rngCol.HorizontalAlignment = xlCenter

                rgt 1
                If (Not t.Summary And Not (t.Duration = 0)) Then
                    If (t.IsPublished) Then
                        m_rngCol = "Y"
                    Else
                        m_rngCol = "N"
                    End If
                End If
                m_rngCol.HorizontalAlignment = xlCenter


                If (m_WoD = "W") Then        ' Show work if requested
                    Set TSVPlannedWork = t.TimeScaleData(dGntStart, dGntFinish, _
                                                         Type:=pjTaskTimescaledWork, TimescaleUnit:=TimescaleUnit)
                    Set TSVActualWork = t.TimeScaleData(dGntStart, dGntFinish, _
                                                        Type:=pjTaskTimescaledActualWork, TimescaleUnit:=TimescaleUnit)

                    For p = 1 To TSVActualWork.Count
                        If Not TSVActualWork(p).Value = "" And Not TSVActualWork(p).Value = 0 Then        ' If there are actuals for that period (p)
                            If TSVActualWork(p).Value = TSVPlannedWork(p).Value Then        ' The plan should be the same as actuals, if so show AC using white font
                                m_rngRow.Offset(0, rngLastCol.Column + p - 1) = TSVActualWork(p).Value / 60
                                m_rngRow.Offset(0, rngLastCol.Column + p - 1).NumberFormat = "0.0"
                                m_rngRow.Offset(0, rngLastCol.Column + p - 1).Font.Color = vbGreen
                            ElseIf Not TSVPlannedWork(p).Value = "" And Not TSVPlannedWork(p).Value = 0 Then        ' If PV <> AC then If there is Planned work for that period (p), show AC using red font
                                m_rngRow.Offset(0, rngLastCol.Column + p - 1) = TSVActualWork(p).Value / 60
                                m_rngRow.Offset(0, rngLastCol.Column + p - 1).NumberFormat = "0.0"
                                m_rngRow.Offset(0, rngLastCol.Column + p - 1).Font.Color = vbYellow
                                m_rngRow.Offset(0, rngLastCol.Column + p - 1).AddComment ("P. W. : " + Format(TSVPlannedWork(p).Value / 60, "0.0"))
                            End If
                        Else        ' There are no actuals for that period (p), therefore show planned work
                            If Not TSVPlannedWork(p).Value = "" And Not TSVPlannedWork(p).Value = 0 Then        ' If there is work planned for that period (p), show it using yellow font
                                m_rngRow.Offset(0, rngLastCol.Column + p - 1) = TSVPlannedWork(p).Value / 60
                                m_rngRow.Offset(0, rngLastCol.Column + p - 1).NumberFormat = "0.0"
                                m_rngRow.Offset(0, rngLastCol.Column + p - 1).Font.Color = vbWhite
                            End If
                        End If
                    Next
                End If


                '------------------------------------------ If Assignment info was requested ----------------------------------
                If (gbIncludeResInfo) Then
                    For Each Asgn In t.Assignments

                        dwn 1
                        wsSheet1.cells(m_rngRow.row, 1) = "A" + CStr(t.OutlineLevel)        ' Indicate it's an Assignment and outlevel
                        wsSheet1.cells(m_rngRow.row, 1).Font.Name = "Kartika"
                        wsSheet1.cells(m_rngRow.row, 1).Font.Size = 8

                        Set m_rngCol = m_rngRow.Offset(0, iColumns + 1)
                        m_rngCol = Asgn.ResourceName
                        m_rngCol.Font.Color = RGB(51, 102, 255)

                        rgt 1
                        If IsNumeric(Asgn.Work) Then m_rngCol = (Asgn.Work / 60)        ' & " hours" 'It was 480?
                        m_rngCol.NumberFormat = "#,##0.0"
                        m_rngCol.Font.Color = RGB(51, 102, 255)

                        rgt 1
                        If IsNumeric(Asgn.BaselineWork) Then m_rngCol = (Asgn.BaselineWork / 60)        ' & " hours" 'It was 480?
                        m_rngCol.NumberFormat = "#,##0.0"
                        m_rngCol.Font.Color = RGB(51, 102, 255)

                        rgt 1
                        If IsNumeric(Asgn.ActualWork) Then m_rngCol = (Asgn.ActualWork / 60)        ' & " hours" 'It was 480?
                        m_rngCol.NumberFormat = "#,##0.0"
                        m_rngCol.Font.Color = RGB(51, 102, 255)

                        rgt 1
                        If IsNumeric(Asgn.RemainingWork) Then m_rngCol = (Asgn.RemainingWork / 60)        ' & " hours" 'It was 480?
                        m_rngCol.NumberFormat = "#,##0.0"
                        m_rngCol.Font.Color = RGB(51, 102, 255)

                        rgt 1
                        If IsDate(Asgn.Start) Then
                            m_rngCol = DateValue(Asgn.Start)        ' (Asgn.ActualWork / 450) & " Days"
                            m_rngCol.Font.Color = RGB(51, 102, 255)
                            m_rngCol.NumberFormat = "mm/dd/yy"
                        End If

                        rgt 1
                        If IsDate(Asgn.Finish) Then
                            m_rngCol = DateValue(Asgn.Finish)
                            m_rngCol.Font.Color = RGB(51, 102, 255)
                            m_rngCol.NumberFormat = "mm/dd/yy"
                        End If

                        rgt 1
                        If IsDate(Asgn.BaselineStart) Then
                            m_rngCol = DateValue(Asgn.BaselineStart)        ' (Asgn.ActualWork / 450) & " Days"
                            m_rngCol.NumberFormat = "mm/dd/yy"
                            m_rngCol.Font.Color = RGB(51, 102, 255)
                        Else
                            m_rngCol = Asgn.BaselineStart        'N/A

                            m_rngCol.HorizontalAlignment = xlRight
                            m_rngCol.Font.Color = RGB(255, 0, 0)        ' Red font
                        End If

                        rgt 1
                        If IsDate(Asgn.BaselineFinish) Then
                            m_rngCol = DateValue(Asgn.BaselineFinish)
                            m_rngCol.NumberFormat = "mm/dd/yy"
                            m_rngCol.Font.Color = RGB(51, 102, 255)
                        Else
                            m_rngCol = Asgn.BaselineFinish        'N/A
                            m_rngCol.HorizontalAlignment = xlRight
                            m_rngCol.Font.Color = RGB(255, 0, 0)        ' Red font
                        End If


                        rgt 1
                        If IsDate(Asgn.ActualStart) Then
                            m_rngCol = DateValue(Asgn.ActualStart)        ' (Asgn.ActualWork / 450) & " Days"
                            m_rngCol.NumberFormat = "mm/dd/yy"
                        Else
                            m_rngCol = Asgn.ActualStart        'N/A
                            m_rngCol.HorizontalAlignment = xlRight
                        End If
                        m_rngCol.Font.Color = RGB(51, 102, 255)

                        rgt 1
                        If IsDate(Asgn.ActualFinish) Then
                            m_rngCol = DateValue(Asgn.ActualFinish)
                            m_rngCol.NumberFormat = "mm/dd/yy"
                        Else
                            m_rngCol = Asgn.ActualFinish        'N/A
                            m_rngCol.HorizontalAlignment = xlRight
                        End If
                        m_rngCol.Font.Color = RGB(51, 102, 255)

                        rgt 1
                        'm_rngCol = Chr$(149) ' Ahead or Behind schedule?
                        'm_rngCol.Font.Size = 11
                        'm_rngCol.Font.Color = RGB(255, 0, 0)
                        'If (IsDate(Asgn.BaselineStart)) Then
                        '    If (IsDate(Asgn.ActualStart)) Then
                        '        If (t.PercentComplete = 100 Or Asgn.ActualStart <= Asgn.BaselineStart) Then
                        '            m_rngCol.Font.Color = RGB(0, 128, 0)
                        '        End If
                        '    Else
                        '        m_rngCol.Font.Color = RGB(0, 128, 0)
                        '    End If
                        'End If

                        'm_rngCol.Font.Bold = True
                        'm_rngCol.HorizontalAlignment = xlCenter
                        rgt 1
                        'm_rngCol = Chr$(149) ' Over or under estimates
                        'm_rngCol.Font.Size = 11
                        'm_rngCol.Font.Color = RGB(255, 0, 0)
                        'If (IsNumeric(Asgn.BaselineWork)) Then
                        '    If (t.PercentComplete = 100 Or Asgn.Work > Asgn.BaselineWork) Then
                        '        m_rngCol.Font.Color = RGB(0, 128, 0)
                        '    End If
                        'End If
                        'm_rngCol.Font.Bold = True
                        'm_rngCol.HorizontalAlignment = xlCenter
                        rgt 1


                        If (m_WoD = "W") Then        ' Show work if requested
                            Set TSVPlannedWork = Asgn.TimeScaleData(dGntStart, dGntFinish, _
                                                                    Type:=pjAssignmentTimescaledWork, TimescaleUnit:=TimescaleUnit)
                            Set TSVActualWork = Asgn.TimeScaleData(dGntStart, dGntFinish, _
                                                                   Type:=pjAssignmentTimescaledActualWork, TimescaleUnit:=TimescaleUnit)        'pjAssignmentTimescaledWork
                            For p = 1 To TSVActualWork.Count
                                If Not TSVActualWork(p).Value = "" And Not TSVActualWork(p).Value = 0 Then        ' If there are actuals for that period (p)
                                    If TSVActualWork(p).Value = TSVPlannedWork(p).Value Then        ' The plan should be the same as actuals, if so show AC using white font
                                        m_rngRow.Offset(0, rngLastCol.Column + p - 1).Font.Size = 8
                                        m_rngRow.Offset(0, rngLastCol.Column + p - 1).Font.Color = vbGreen
                                        m_rngRow.Offset(0, rngLastCol.Column + p - 1) = TSVActualWork(p).Value / 60
                                        m_rngRow.Offset(0, rngLastCol.Column + p - 1).NumberFormat = "0.0"
                                    ElseIf Not TSVPlannedWork(p).Value = "" And Not TSVPlannedWork(p).Value = 0 Then        ' If PV <> AC then If there is Planned work for that period (p), , show AC using red font
                                        m_rngRow.Offset(0, rngLastCol.Column + p - 1) = TSVActualWork(p).Value / 60
                                        m_rngRow.Offset(0, rngLastCol.Column + p - 1).NumberFormat = "0.0"
                                        m_rngRow.Offset(0, rngLastCol.Column + p - 1).Font.Color = vbYellow
                                        m_rngRow.Offset(0, rngLastCol.Column + p - 2).AddComment ("P. W. : " + Format(TSVPlannedWork(p).Value / 60, "0.0"))
                                    End If
                                Else        ' There are no actuals for that period (p), therefore show planned work
                                    If Not TSVPlannedWork(p).Value = "" And Not TSVPlannedWork(p).Value = 0 Then        ' If there is work planned for that period (p), show it using yellow font
                                        m_rngRow.Offset(0, rngLastCol.Column + p - 1) = TSVPlannedWork(p).Value / 60
                                        m_rngRow.Offset(0, rngLastCol.Column + p - 1).NumberFormat = "0.0"
                                        m_rngRow.Offset(0, rngLastCol.Column + p - 1).Font.Color = vbWhite
                                    End If
                                End If
                            Next
                        End If
                    Next
                End If
                If (t.PercentComplete = 100) Then
                    wsSheet1.Range(wsSheet1.cells(rngTasKRow.row, 1), wsSheet1.cells(m_rngCol.row, m_rngCol.Column)).Interior.Color = RGB(224, 224, 224)
                ElseIf (t.PercentComplete > 0) Then        ' In progress '(Not t.Summary And t.PercentComplete = 0 And t.Start < Now()) Then
                    wsSheet1.Range(wsSheet1.cells(rngTasKRow.row, 1), wsSheet1.cells(rngTasKRow.row, m_rngCol.Column)).Interior.Color = RGB(182, 244, 182)
                ElseIf (t.Start > Now() And t.Start < Now() + 14) Then        ' Starting in the next 14 days
                    wsSheet1.Range(wsSheet1.cells(rngTasKRow.row, 1), wsSheet1.cells(rngTasKRow.row, m_rngCol.Column)).Interior.Color = RGB(255, 255, 155)
                End If
                iTotTasks = iTotTasks + 1
            End If
        End If
    Next t

    '''''''''''''''''''''''''''''''''''''''''''''''''''''' Gantt Chart logic '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' m_rngRow contains the last row
    ' rngLastCol contains the last col before the gantt (at the title level)

    Dim s, K, l, M, sCellGantt, sGanttCoord, sCellStart, sCellFinish, sStartCoord, sFinishCoord As String
    Dim vArr As Variant







    ' Start coordenates for the Gantt Chart
    iX1GanttCoord = rngLastCol.Offset(1, 1).row
    iY1GanttCoord = rngLastCol.Offset(1, 1).Column


    ' Finish coordinates for the Gantt Chart
    iX2GanttCoord = m_rngRow.Offset(0, 1).row
    iY2GanttCoord = m_rngRow.Offset(0, iY1GanttCoord + iNumOfWeeksInGanttChart - 2).Column

    sCellGantt = "R" + CStr(iX1GanttCoord) + "C" + CStr(iY1GanttCoord)
    sCellStart = "R" + CStr(iX1GanttCoord) + "C" + CStr(iY1GanttCoord - 10)
    sCellFinish = "R" + CStr(iX1GanttCoord) + "C" + CStr(iY1GanttCoord - 9)

    sGanttCoord = xlApp.ConvertFormula(sCellGantt, xlR1C1LB, xlA1LB)
    sStartCoord = xlApp.ConvertFormula(sCellStart, xlR1C1LB, xlA1LB)
    sFinishCoord = xlApp.ConvertFormula(sCellFinish, xlR1C1LB, xlA1LB)

    vArr = Split(sGanttCoord, "$")
    M = vArr(1)        'Column where Gantt starts

    vArr = Split(sStartCoord, "$")
    K = vArr(1)        'Start column

    vArr = Split(sFinishCoord, "$")
    l = vArr(1)        ' Finish column

    ' s = "=IF(AND(M$3>=$K4,M$3-5<$L4),
    '          IF($K4>M$3-5,
    '                 IF(M$3>$L4,
    '                      CONCATENATE(WEEKDAY($K4,2),NETWORKDAYS($K4,$L4)),
    '                      CONCATENATE(WEEKDAY($K4,2),NETWORKDAYS($K4,M$3))),
    '                 IF(M$3>$L4,
    '                      CONCATENATE(1,NETWORKDAYS(M$3-4,$L4)),
    '                      CONCATENATE(1,NETWORKDAYS(M$3-4,M$3)))
    '           ),
    '           """")"
    '

    s = "=IF(AND(" + M + "$3>=$" + K + "4," + M + "$3-5<$" + l + "4),IF($" + K + "4>" + M + "$3-5,IF(" + M + _
        "$3>$" + l + "4,CONCATENATE(WEEKDAY($" + K + "4,2),NETWORKDAYS($" + K + "4,$" + l + "4)),CONCATENATE(WEEKDAY($" + K + _
        "4,2),NETWORKDAYS($" + K + "4," + M + "$3))),IF(" + M + "$3>$" + l + "4,CONCATENATE(1,NETWORKDAYS(" + _
        M + "$3-5,$" + l + "4)),CONCATENATE(1,NETWORKDAYS(" + M + "$3-5," + M + "$3)))),"""")"

    With wsSheet1
        .Range(.cells(iX1GanttCoord, iY1GanttCoord), .cells(iX2GanttCoord, iY2GanttCoord)).Interior.Color = RGB(204, 255, 255)
        .Range(M + "4").Select
        With .Range(.cells(iX1GanttCoord, iY1GanttCoord), .cells(iX2GanttCoord, iY2GanttCoord))
            .Borders(xlEdgeLeftLB).LineStyle = xlContinuousLB
            .Borders(xlEdgeLeftLB).Color = RGB(153, 204, 255)
            .Borders(xlEdgeRightLB).LineStyle = xlContinuousLB
            .Borders(xlEdgeRightLB).Color = RGB(153, 204, 255)
            .Borders(xlEdgeTopLB).LineStyle = xlContinuousLB
            .Borders(xlEdgeTopLB).Color = RGB(153, 204, 255)
            .Borders(xlEdgeBottomLB).LineStyle = xlContinuousLB
            .Borders(xlEdgeBottomLB).Color = RGB(153, 204, 255)
            .Borders(xlInsideVerticalLB).LineStyle = xlContinuousLB
            .Borders(xlInsideVerticalLB).Weight = xlThin
            .Borders(xlInsideVerticalLB).Color = RGB(153, 204, 255)
            .Borders(xlInsideHorizontalLB).LineStyle = xlContinuousLB
            .Borders(xlInsideHorizontalLB).Weight = xlThin
            .Borders(xlInsideHorizontalLB).Color = RGB(153, 204, 255)

            If (m_WoD = "D") Then        ' Show days if requested
                .Formula = s
            End If

            ' .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(M$3>=$K4,M$3-5<$L4, $A4="T")"
            .FormatConditions.Add Type:=xlExpressionLB, Formula1:="=AND(" + M + "$3>=$" + K + "4," + M + "$3-5<$" + l + "4,$A4=""M"")"
            .FormatConditions.Add Type:=xlExpressionLB, Formula1:="=AND(" + M + "$3>=$" + K + "4," + M + "$3-5<$" + l + "4,left($A4,1)=""S"")"
            .FormatConditions.Add Type:=xlExpressionLB, Formula1:="=AND(" + M + "$3>=$" + K + "4," + M + "$3-5<$" + l + "4,left($A4,1)=""T"")"
            .FormatConditions.Add Type:=xlExpressionLB, Formula1:="=AND(" + M + "$3>=$" + K + "4," + M + "$3-5<$" + l + "4,left($A4,1)=""A"")"

            With .FormatConditions(1)        ' Milestone color
                .SetFirstPriority
                .Interior.Color = RGB(255, 50, 50)
                ''' VOLVER '''' .Font.Color = RGB(255, 255, 255)
            End With

            With .FormatConditions(2)        ' Summary Task color
                .Interior.Color = RGB(0, 0, 0)
                ''' VOLVER '''' .Font.Color = RGB(204, 255, 255)
            End With

            With .FormatConditions(3)        ' Task color
                .Interior.Color = RGB(37, 64, 97)        ' 88, 88, 88 / 39, 136, 120
                ''' VOLVER '''' .Font.Color = RGB(0, 0, 0)
            End With

            With .FormatConditions(4)        ' Assignment color
                .Interior.Color = RGB(149, 55, 53)        '139, 136, 120 / 162, 181, 205
                ''' VOLVER '''' .Font.Color = RGB(0, 0, 0)
            End With
        End With
    End With


    ' Painting columns (Gray for past, and different colors for even and odd months)
    With wsSheet1
        For i = iY1GanttCoord To iY2GanttCoord

            If (m_WoD = "W") Then        ' Reduce font if I need to show Work
                .Range(.cells(4, i), .cells(iX2GanttCoord, i)).Font.Size = 8
            End If

            If (.cells(3, i) < Now()) Then
                .Range(.cells(1, i), .cells(iX2GanttCoord, i)).Interior.Color = RGB(224, 224, 224)
                With .Range(.cells(4, i), .cells(iX2GanttCoord, i))
                    .Borders(xlEdgeLeftLB).Color = RGB(150, 150, 150)
                    .Borders(xlEdgeRightLB).Color = RGB(150, 150, 150)
                    .Borders(xlEdgeTopLB).Color = RGB(150, 150, 150)
                    .Borders(xlEdgeBottomLB).Color = RGB(150, 150, 150)
                    .Borders(xlInsideVerticalLB).Weight = xlThin
                    .Borders(xlInsideVerticalLB).Color = RGB(150, 150, 150)
                    .Borders(xlInsideHorizontalLB).Weight = xlThin
                    .Borders(xlInsideHorizontalLB).Color = RGB(150, 150, 150)
                End With
            ElseIf (Month(.cells(2, i)) Mod 2) = 0 Then
                .Range(.cells(1, i), .cells(iX2GanttCoord, i)).Interior.Color = RGB(122, 197, 205)
                With .Range(.cells(iX1GanttCoord, i), .cells(iX2GanttCoord, i))
                    .Borders(xlEdgeLeftLB).Color = RGB(95, 158, 160)
                    .Borders(xlEdgeRightLB).Color = RGB(95, 158, 160)
                    .Borders(xlEdgeTopLB).Color = RGB(95, 158, 160)
                    .Borders(xlEdgeBottomLB).Color = RGB(95, 158, 160)
                    .Borders(xlInsideVerticalLB).Weight = xlThin
                    .Borders(xlInsideVerticalLB).Color = RGB(95, 158, 160)
                    .Borders(xlInsideHorizontalLB).Weight = xlThin
                    .Borders(xlInsideHorizontalLB).Color = RGB(95, 158, 160)
                End With
            End If
        Next
    End With

    wsSheet1.Range("W4").Select

    '''pj.ProjectSummaryTask.OutlineHideSubTasks

    Call tideUp


    oProgress.Hide

    xlApp.ScreenUpdating = True
    xlApp.DisplayAlerts = True
    xlApp.Calculation = xlCalculationAutomatic

    '''''''''''''''''''''''''''''''''''''''''''''''''''End of New Code'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    '''AppActivate "Microsoft Project"


    Call MsgBox("Task Hierarchy complete (" & iTotTasks & " Tasks added)", vbInformation Or vbOKOnly)



    xlApp.Visible = True




    AppActivate wbWkBk1.Name

    wbWkBk1.Application.WindowState = xlMaximizedLB

    xlApp.ActiveWindow.FreezePanes = True

    ''''frmToolbox.Show False


    Set xlApp = Nothing
End Sub





Function calculateTotals()






