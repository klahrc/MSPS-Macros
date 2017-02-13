VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressDialogue 
   Caption         =   "Progress"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6135
   OleObjectBlob   =   "ProgressDialogue.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressDialogue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Implements IPRogressBar

Private Cancelled As Boolean, showTime As Boolean, showTimeLeft As Boolean
Private Value_Minimum As Long, Value_Maximum As Long, Value_Current As Long
Dim startTime As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub IProgressBar_Show()
    Me.Show
End Sub
Public Sub IProgressBar_Hide()
    Me.Hide
End Sub

'Title will be the title of the dialogue.
'Status will be the label above the progress bar, and can be changed with SetStatus.
'Min is the progress bar minimum value, only set by calling configure.
'Max is the progress bar maximum value, only set by calling configure.
'CancelButtonText is the caption of the cancel button. If set to vbNullString, it is hidden.
'optShowTimeElapsed controls whether the progress bar computes and displays the time elapsed.
'optShowTimeRemaining controls whether the progress bar estimates and displays the time remaining.
'calling Configure sets the current value equal to Min.
'calling Configure resets the current run time.
Public Sub IProgressBar_Configure(ByVal title As String, ByVal Status As String, _
                     ByVal Min As Long, ByVal Max As Long, _
                     Optional ByVal CancelButtonText As String = "Cancel", _
                     Optional ByVal optShowTimeElapsed As Boolean = True, _
                     Optional ByVal optShowTimeRemaining As Boolean = True)
    Me.Caption = title
    lblStatus.Caption = Status
    Value_Minimum = Min
    Value_Maximum = Max
    Value_Current = Min
    CancelButton.Visible = Not CancelButtonText = vbNullString
    CancelButton.Caption = CancelButtonText
    startTime = GetTickCount
    showTime = optShowTimeElapsed
    showTimeLeft = optShowTimeRemaining
    lblRunTime.Caption = ""
    lblRemainingTime.Caption = ""
    Cancelled = False
End Sub

'Set the label text above the status bar
Public Property Let IProgressBar_Status(ByVal new_status As String)
    lblStatus.Caption = new_status
    DoEvents
End Property

'Set the value of the status bar, a long which is snapped to a value between Min and Max
Public Property Let IProgressBar_Value(ByVal new_value As Long)
    If new_value < Value_Minimum Then new_value = Value_Minimum
    If new_value > Value_Maximum Then new_value = Value_Maximum
    Dim progress As Double, RunTime As Long
    Value_Current = new_value
    progress = (Value_Current - Value_Minimum) / (Value_Maximum - Value_Minimum)
    ProgressBar.Width = 292 * progress
    lblPercent = Int(progress * 10000) / 100 & "%"
    RunTime = IProgressBar_RunTime()
    If showTime Then lblRunTime.Caption = "Time Elapsed: " & GetRunTimeString(RunTime, True)
    If showTimeLeft And progress > 0 Then _
        lblRemainingTime.Caption = "Est. Time Left: " & GetRunTimeString(RunTime * (1 - progress) / progress, False)
    DoEvents
End Property

'Get the time (in milliseconds) since the progress bar "Configure" routine was last called
Public Property Get IProgressBar_RunTime() As Long
    IProgressBar_RunTime = GetTickCount - startTime
End Property

'Get the time (in hours, minutes, seconds) since "Configure" was last called
Public Property Get IProgressBar_FormattedRunTime() As String
    IProgressBar_FormattedRunTime = GetRunTimeString(GetTickCount - startTime)
End Property

'Formats a time in milliseconds as hours, minutes, seconds.milliseconds
'Milliseconds are excluded if showMsecs is set to false
Private Function GetRunTimeString(ByVal RunTime As Long, Optional ByVal showMsecs As Boolean = True) As String
    Dim msecs&, hrs&, mins&, secs#
    msecs = RunTime
    hrs = Int(msecs / 3600000)
    mins = Int(msecs / 60000) - 60 * hrs
    secs = msecs / 1000 - 60 * (mins + 60 * hrs)
    GetRunTimeString = IIf(hrs > 0, hrs & " hours ", "") _
                     & IIf(mins > 0, mins & " minutes ", "") _
                     & IIf(secs > 0, IIf(showMsecs, secs, Int(secs + 0.5)) & " seconds", "")
End Function

'Returns the current value of the progress bar
Public Property Get IProgressBar_Value() As Long
    IProgressBar_Value = Value_Current
End Property

'Returns whether or not the cancel button has been pressed.
'The ProgressDialogue must be polled regularily to detect whether cancel was pressed.
Public Property Get IProgressBar_cancelIsPressed() As Boolean
    IProgressBar_cancelIsPressed = Cancelled
End Property

'Recalls that cancel was pressed so that they calling routine can be notified next time it asks.
Private Sub CancelButton_Click()
    Cancelled = True
    lblStatus.Caption = "Cancelled By User. Please Wait."
End Sub


