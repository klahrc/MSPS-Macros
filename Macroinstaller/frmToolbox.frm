VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmToolbox 
   Caption         =   "MSPS Toolbox"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   1905
   OleObjectBlob   =   "frmToolbox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmToolbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' DeleteEmptyTasks

Private Sub lblAddTaskPrefix_Click()
    Call MI_AddTP
End Sub

Private Sub lblAssignment_Click()
    Call MI_Asgn
End Sub
Private Sub lblASAP_Click()
Call MI_Chg2ASAP
End Sub

Private Sub lblResPlan_Click()
    Call MI_ResPlan
End Sub

Private Sub lblTH_Click()
    Call MI_TH
End Sub
