VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTasksTHChoices 
   Caption         =   "Please choose..."
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4920
   OleObjectBlob   =   "frmTasksTHChoices.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTasksTHChoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkResInfo_Click()
    gbIncludeResInfo = chkResInfo.Value
    chkResInfo.Font.Bold = chkResInfo.Value
End Sub

Private Sub imgOK_Click()
    Me.Hide
    frmTH.Show False
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
    chkResInfo.Value = True
    optAllTasks.Value = True

    gbIncludeResInfo = True
End Sub
