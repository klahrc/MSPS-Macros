VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Task Hierarchy Report - Settings"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8565
   OleObjectBlob   =   "TH.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAB_Click()
    chkAB.Font.Bold = chkAB.Value
End Sub

Private Sub chkAllActuals_Click()
    chkActualWork.Value = chkAllActuals.Value
    chkActualDates.Value = chkAllActuals.Value
    chkActualCost.Value = chkAllActuals.Value
End Sub

Private Sub chkAllAssessment_Click()
    chkAB.Value = chkAllAssessment.Value
    chkOU.Value = chkAllAssessment.Value
End Sub

Private Sub chkAllBaseline_Click()
    chkBaselineWork.Value = chkAllBaseline.Value
    chkBaselineDates.Value = chkAllBaseline.Value
    chkBaselineCost.Value = chkAllBaseline.Value
End Sub

Private Sub chkAllCompletion_Click()
    chkCompletionStatus.Value = chkAllCompletion.Value
    chkPublished.Value = chkAllCompletion.Value
End Sub

Private Sub chkAllGantt_Click()
    chkGantt.Value = chkAllGantt.Value
    chkTimePhaseData.Value = chkAllGantt.Value
End Sub

Private Sub chkAllPlanned_Click()
    chkPlannedWork.Value = chkAllPlanned.Value
    chkPlannedDates.Value = chkAllPlanned.Value
    chkPlannedCost.Value = chkAllPlanned.Value

End Sub

Private Sub chkAllRemaining_Click()
    chkRemainingWork.Value = chkAllRemaining.Value
    chkRemainingDays.Value = chkAllRemaining.Value
    chkRemainingCost.Value = chkAllRemaining.Value
End Sub

'Work
Private Sub chkAllWork_Click()
    chkActualWork.Value = chkAllWork.Value
    chkBaselineWork.Value = chkAllWork.Value
    chkPlannedWork.Value = chkAllWork.Value
    chkRemainingWork.Value = chkAllWork.Value
End Sub

Private Sub chkActualWork_Click()
    chkActualWork.Font.Bold = chkActualWork.Value
End Sub

Private Sub chkBaselineWork_Click()
    chkBaselineWork.Font.Bold = chkBaselineWork.Value
End Sub

Private Sub chkCompletionStatus_Click()
    chkCompletionStatus.Font.Bold = chkCompletionStatus.Value
End Sub

Private Sub chkGantt_Click()
    chkGantt.Font.Bold = chkGantt.Value
End Sub

Private Sub chkkMore_Click()
    chkGantt.Value = chkkMore.Value
    chkTimePhaseData.Value = chkkMore.Value
    chkCompletionStatus.Value = chkkMore.Value
    chkPublished.Value = chkkMore.Value
    chkAB.Value = chkkMore.Value
    chkOU.Value = chkkMore.Value
End Sub

Private Sub chkOU_Click()
    chkOU.Font.Bold = chkOU.Value
End Sub

Private Sub chkPlannedWork_Click()
    chkPlannedWork.Font.Bold = chkPlannedWork.Value
End Sub

Private Sub chkPublished_Click()
    chkPublished.Font.Bold = chkPublished.Value
End Sub

Private Sub chkRemainingWork_Click()
    chkRemainingWork.Font.Bold = chkRemainingWork.Value
End Sub

'Time
Private Sub chkAllDates_Click()
    chkActualDates.Value = chkAllDates.Value
    chkBaselineDates.Value = chkAllDates.Value
    chkPlannedDates.Value = chkAllDates.Value
    chkRemainingDays.Value = chkAllDates.Value
End Sub
Private Sub chkActualDates_Click()
    chkActualDates.Font.Bold = chkActualDates.Value
End Sub

Private Sub chkBaselineDates_Click()
    chkBaselineDates.Font.Bold = chkBaselineDates.Value
End Sub

Private Sub chkPlannedDates_Click()
    chkPlannedDates.Font.Bold = chkPlannedDates.Value
End Sub

Private Sub chkRemainingDays_Click()
    chkRemainingDays.Font.Bold = chkRemainingDays.Value
End Sub

'Cost
Private Sub chkAllCost_Click()
    chkActualCost.Value = chkAllCost.Value
    chkBaselineCost.Value = chkAllCost.Value
    chkPlannedCost.Value = chkAllCost.Value
    chkRemainingCost.Value = chkAllCost.Value
End Sub
Private Sub chkActualCost_Click()
    chkActualCost.Font.Bold = chkActualCost.Value
End Sub

Private Sub chkBaselineCost_Click()
    chkBaselineCost.Font.Bold = chkBaselineCost.Value
End Sub

Private Sub chkPlannedCost_Click()
    chkPlannedCost.Font.Bold = chkPlannedCost.Value
End Sub

Private Sub chkRemainingCost_Click()
    chkRemainingCost.Font.Bold = chkRemainingCost.Value
End Sub

Private Sub chkTimePhaseData_Click()
    chkTimePhaseData.Font.Bold = chkTimePhaseData.Value
End Sub


Private Sub Image5_Click()

End Sub
