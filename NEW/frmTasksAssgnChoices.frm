VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTasksAssgnChoices 
   Caption         =   "Please choose..."
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4920
   OleObjectBlob   =   "frmTasksAssgnChoices.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTasksAssgnChoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkTaskInfo_Click()
    gbIncludeTaskInfo = chkTaskInfo.Value
    chkTaskInfo.Font.Bold = chkTaskInfo.Value
End Sub

Private Sub imgOK_Click()
    Me.Hide
    frmAssgn.Show False
End Sub

Private Sub optAllTasks_Click()
    gbProcessAllTasks = True
    optAllTasks.Font.Bold = True
    optSelectedTasks.Font.Bold = False
End Sub

Private Sub optSelectedTasks_Click()
    gbProcessAllTasks = False
    ''''optSelectedTasks.Font.Bold = True '''(DISABLED BY NOW)
    ''''optAllTasks.Font.Bold = False
    MsgBox "This option is disabled", vbOKOnly Or vbInformation, "ATTENTION"
    optAllTasks.Value = True
    

End Sub

Private Sub UserForm_Initialize()
    chkTaskInfo.Value = True
    optAllTasks.Value = True

    gbIncludeTaskInfo = True
End Sub
