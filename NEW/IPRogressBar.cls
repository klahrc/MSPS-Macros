VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPRogressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Title will be the title of the dialogue.
'Status will be the label above the progress bar, and can be changed with SetStatus.
'Min is the progress bar minimum value, only set by calling configure.
'Max is the progress bar maximum value, only set by calling configure.
'CancelButtonText is the caption of the cancel button. If set to vbNullString, it is hidden.
'optShowTimeElapsed controls whether the progress bar computes and displays the time elapsed.
'optShowTimeRemaining controls whether the progress bar estimates and displays the time remaining.
'calling Configure sets the current value equal to Min.
'calling Configure resets the current run time.
Public Sub Configure(ByVal title As String, ByVal Status As String, _
                     ByVal Min As Long, ByVal Max As Long, _
                     Optional ByVal CancelButtonText As String = "Cancel", _
                     Optional ByVal optShowTimeElapsed As Boolean = True, _
                     Optional ByVal optShowTimeRemaining As Boolean = True)

End Sub

'Set the label text above the status bar
Public Property Let Status(ByVal new_status As String)

End Property

'Set the value of the status bar, a long which is snapped to a value between Min and Max
Public Property Let Value(ByVal new_value As Long)

End Property

'Get the time (in milliseconds) since the progress bar "Configure" routine was last called
Public Property Get RunTime() As Long

End Property

'Get the time (in hours, minutes, seconds) since "Configure" was last called
Public Property Get FormattedRunTime() As String

End Property

'Returns the current value of the progress bar
Public Property Get Value() As Long

End Property

'Returns whether or not the cancel button has been pressed.
'The ProgressDialogue must be polled regularily to detect whether cancel was pressed.
Public Property Get cancelIsPressed() As Boolean

End Property

Public Sub Show()

End Sub
Public Sub Hide()

End Sub





